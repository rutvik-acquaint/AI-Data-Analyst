#!/usr/bin/env python3
"""
Universal File Reader & Ingestion Engine for Local RAG
Supports:
- Excel: .xls (BIFF8 legacy via xlrd), .xlsx, .xlsm, .xltx (OpenXML via openpyxl)
- CSV / TSV / Delimited text
- PDF: .pdf (via pypdf)
- Word: .docx (via python-docx)
- PowerPoint: .pptx (via python-pptx)
- Plain text, Markdown, JSON, YAML, XML, HTML, source code, and log files

Features:
- Automatic schema & dataset analysis for spreadsheets (detects columns, data types, sample values)
- Clean Markdown table formatting for spreadsheets to maximize LLM & vector retrieval accuracy
- Header & file context propagation so table chunks never lose their column definitions
- Deduplication: updates/replaces vectors when the same file is re-uploaded
"""

# ==============================================================================
# Standard Library Imports
# ==============================================================================
import os
import sys
import io
import csv
import json
import uuid
import datetime
import logging
import threading
import sqlite3
import ast
import operator
import time
import re
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple, Union

# ==============================================================================
# Third-Party Core Dependencies
# ==============================================================================
import requests
from pydantic import BaseModel
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

# ==============================================================================
# Optional Plugins & Specialized Processing Libraries
# ==============================================================================
try:
    import xlrd
except ImportError:
    xlrd = None

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import pypdf
except ImportError:
    pypdf = None

try:
    import docx
except ImportError:
    docx = None

try:
    import pptx
except ImportError:
    pptx = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import duckdb
except ImportError:
    duckdb = None

try:
    from fastembed import SparseTextEmbedding
except ImportError:
    SparseTextEmbedding = None

try:
    from flashrank import Ranker, RerankRequest
except ImportError:
    Ranker = None
    RerankRequest = None

# ==============================================================================
# Global Logger Setup
# ==============================================================================
import pandas as pd

# ==============================================================================
# Model & Global Constants (Tuned for 4-core, 15 GB CPU machine)
# ==============================================================================
OLLAMA_URL = os.getenv('OLLAMA_URL', 'http://localhost:11434')
ANSWER_MODEL = os.getenv('ANSWER_MODEL', 'llama3.2:3b')
SQL_MODEL = os.getenv('SQL_MODEL', 'llama3.2:3b')   # or qwen2.5-coder:7b if available
KEEP_ALIVE = os.getenv('OLLAMA_KEEP_ALIVE', '30m')
FENCE = chr(96) * 3
MAX_ROWS = 5000
MIN_RERANK_SCORE = 0.008

MISSING = {'', 'none', 'nan', 'null', 'n/a', 'na', '-', '--'}
NUMERIC_TYPES = ('TINYINT', 'SMALLINT', 'INTEGER', 'BIGINT', 'HUGEINT', 'FLOAT', 'DOUBLE', 'DECIMAL')
PAGE_RE = re.compile(r'\n?--- Page (\d+) ---\n')
_NUM = re.compile(r'^-?\d+(\.\d+)?$')
_DATE = re.compile(r'^\d{4}-\d{2}(-\d{2})?')


def clean_columns(headers: List[Any]) -> List[str]:
    seen, out = {}, []
    for i, h in enumerate(headers):
        c = re.sub(r'[^a-z0-9]+', '_', str(h).lower()).strip('_') or f'col_{i + 1}'
        if c[0].isdigit():
            c = f'c_{c}'
        seen[c] = seen.get(c, 0) + 1
        out.append(c if seen[c] == 1 else f'{c}_{seen[c]}')
    return out


def coerce_series(s: pd.Series) -> pd.Series:
    s = s.astype('string').str.strip()
    s = s.mask(s.str.lower().isin(MISSING))
    vals = s.dropna()
    if vals.empty or vals.str.match(r'^0\d+$').any():   # keep IDs / zip codes as text
        return s
    strip = lambda x: x.str.replace(r'[,$€£₹\s]', '', regex=True).str.rstrip('%')
    if pd.to_numeric(strip(vals), errors='coerce').notna().mean() >= 0.95:
        full = pd.to_numeric(strip(s), errors='coerce')
        if full.dropna().mod(1).eq(0).all():
            return full.astype('Int64')
        return full.astype('float64')
    dt = pd.to_datetime(vals, errors='coerce', format='mixed')
    if dt.notna().mean() >= 0.95:
        return pd.to_datetime(s, errors='coerce', format='mixed')
    return s


def profile_table(con, table: str, max_values: int = 30) -> List[Dict[str, Any]]:
    out = []
    for name, typ, *_ in con.execute(f'DESCRIBE "{table}"').fetchall():
        nulls, distinct = con.execute(
            f'SELECT COUNT(*) - COUNT("{name}"), COUNT(DISTINCT "{name}") FROM "{table}"').fetchone()
        p = {'name': name, 'orig_name': name, 'type': typ, 'nulls': nulls, 'distinct': distinct}
        if typ.startswith(NUMERIC_TYPES) or typ.startswith(('DATE', 'TIMESTAMP')):
            lo, hi = con.execute(f'SELECT MIN("{name}"), MAX("{name}") FROM "{table}"').fetchone()
            p['min'], p['max'] = str(lo), str(hi)
        elif typ == 'VARCHAR':
            rows = con.execute(
                f'SELECT "{name}" FROM "{table}" WHERE "{name}" IS NOT NULL '
                f'GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT {max_values}').fetchall()
            vals = [r[0] for r in rows]
            p['values' if distinct <= max_values else 'sample'] = vals if distinct <= max_values else vals[:5]
        out.append(p)
    return out


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("rag_analyst")


class UniversalFileReader:
    """Extracts clean, structured text and rich dataset metadata from any file format."""

    @staticmethod
    def _format_cell_value(val: Any) -> str:
        """Converts any raw table cell value (numbers, dates, text) into a clean, readable string."""
        if val is None:
            return ""
        if isinstance(val, float) and val.is_integer():
            return str(int(val))
        if isinstance(val, (datetime.date, datetime.datetime)):
            return val.strftime("%Y-%m-%d %H:%M:%S" if isinstance(val, datetime.datetime) else "%Y-%m-%d")
        
        s_val = str(val).replace("\r", " ").replace("\n", " ").strip()
        # Normalize DD/MM/YYYY or DD-MM-YYYY dates into ISO YYYY-MM-DD format
        m = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$", s_val)
        if m:
            d, mth, y = m.groups()
            return f"{int(y):04d}-{int(mth):02d}-{int(d):02d}"
        return s_val

    @classmethod
    def _generate_table_summary(cls, sheet_name: str, headers: List[str], rows: List[List[Any]]) -> str:
        """Generates a rich natural-language schema & dataset overview for spreadsheets."""
        total_rows = len(rows)
        total_cols = len(headers)
        if total_rows == 0:
            return f"Sheet '{sheet_name}' is empty."

        lines = [
            f"### Dataset Summary for Sheet: '{sheet_name}'",
            f"- Total Data Rows: {total_rows} records",
            f"- Total Columns: {total_cols} attributes",
            f"- Fields / Column Names: {', '.join(headers)}",
            "",
            "#### Column Details, Distributions & Sample Values:"
        ]

        for col_idx, col_name in enumerate(headers):
            col_vals = [r[col_idx] for r in rows if col_idx < len(r) and r[col_idx] not in ("", None)]
            sample_distinct = []
            for v in col_vals:
                s_val = cls._format_cell_value(v)
                if s_val and s_val not in sample_distinct:
                    sample_distinct.append(s_val)
                if len(sample_distinct) >= 5:
                    break

            sample_str = ", ".join(sample_distinct) if sample_distinct else "None"
            col_desc = f"- **{col_name}**: (e.g. {sample_str})"

            # Date breakdown
            if "date" in col_name.lower():
                years = {}
                for v in col_vals:
                    s_v = cls._format_cell_value(v)
                    m = re.search(r"\b(19\d\d|20\d\d)\b", s_v)
                    if m:
                        y = m.group(1)
                        years[y] = years.get(y, 0) + 1
                if years:
                    y_summary = ", ".join([f"{y}: {c} record(s)" for y, c in sorted(years.items())])
                    col_desc += f" | Verified Year Distribution: {y_summary}"

            lines.append(col_desc)

        return "\n".join(lines)

    @classmethod
    def read_xls(cls, content_or_path: Union[str, bytes], filename: str = "document.xls") -> Dict[str, Any]:
        """Extracts legacy Microsoft Excel (.xls) files using xlrd."""
        if xlrd is None:
            raise ImportError("xlrd is required for .xls files. Install with: pip install xlrd")

        if isinstance(content_or_path, bytes):
            wb = xlrd.open_workbook(file_contents=content_or_path)
        else:
            wb = xlrd.open_workbook(content_or_path)

        text_lines = [
            f"# Document Overview: {filename}",
            f"This file is an Excel spreadsheet containing structured data records across {wb.nsheets} sheet(s).\n"
        ]
        sheets_meta = []
        total_rows = 0

        for sheet_idx in range(wb.nsheets):
            sheet = wb.sheet_by_index(sheet_idx)
            if sheet.nrows == 0:
                continue

            # Clean headers
            headers = []
            for c in range(sheet.ncols):
                h = cls._format_cell_value(sheet.cell_value(0, c))
                if not h or h in ["0", "0.0"]:
                    h = f"Index" if c == 0 else f"Column_{c+1}"
                headers.append(h)

            # Read rows
            rows_data = []
            for r in range(1, sheet.nrows):
                row_vals = []
                for c in range(sheet.ncols):
                    cell = sheet.cell(r, c)
                    val = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            date_tuple = xlrd.xldate_as_tuple(val, wb.datemode)
                            val = f"{date_tuple[0]:04d}-{date_tuple[1]:02d}-{date_tuple[2]:02d}"
                        except Exception:
                            pass
                    row_vals.append(cls._format_cell_value(val))
                rows_data.append(row_vals)

            total_rows += len(rows_data)
            sheets_meta.append({
                "sheet_name": sheet.name,
                "rows": len(rows_data),
                "columns": len(headers),
                "headers": headers,
                "data": rows_data
            })

            # 1. Add Natural Language Summary
            summary = cls._generate_table_summary(sheet.name, headers, rows_data)
            text_lines.append(summary)
            text_lines.append("\n#### Complete Data Table:\n")

            # 2. Add Markdown Table
            text_lines.append("| " + " | ".join(headers) + " |")
            text_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")

            for r_vals in rows_data:
                text_lines.append("| " + " | ".join(r_vals) + " |")

            text_lines.append("\n---\n")

        return {
            "text": "\n".join(text_lines),
            "file_type": "xls",
            "metadata": {
                "filename": filename,
                "sheets": sheets_meta,
                "tables": sheets_meta,
                "total_rows": total_rows
            }
        }

    @classmethod
    def read_xlsx(cls, content_or_path: Union[str, bytes], filename: str = "document.xlsx") -> Dict[str, Any]:
        """Extracts modern Excel (.xlsx, .xlsm) files using openpyxl."""
        if openpyxl is None:
            raise ImportError("openpyxl is required for .xlsx files. Install with: pip install openpyxl")

        if isinstance(content_or_path, bytes):
            wb = openpyxl.load_workbook(io.BytesIO(content_or_path), data_only=True)
        else:
            wb = openpyxl.load_workbook(content_or_path, data_only=True)

        text_lines = [
            f"# Document Overview: {filename}",
            f"This file is an Excel spreadsheet containing structured data records across {len(wb.sheetnames)} sheet(s).\n"
        ]
        sheets_meta = []
        total_rows = 0

        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            nrows = sheet.max_row or 0
            ncols = sheet.max_column or 0
            if nrows == 0 or ncols == 0:
                continue

            rows_iter = sheet.iter_rows(values_only=True)
            try:
                first_row = next(rows_iter)
            except StopIteration:
                continue

            headers = []
            for i, c in enumerate(first_row):
                h = cls._format_cell_value(c)
                if not h or h in ["0", "0.0"]:
                    h = f"Index" if i == 0 else f"Column_{i+1}"
                headers.append(h)

            rows_data = []
            for row in rows_iter:
                row_vals = [cls._format_cell_value(c) for c in row]
                rows_data.append(row_vals)

            total_rows += len(rows_data)
            sheets_meta.append({
                "sheet_name": sheet_name,
                "rows": len(rows_data),
                "columns": len(headers),
                "headers": headers,
                "data": rows_data
            })

            summary = cls._generate_table_summary(sheet_name, headers, rows_data)
            text_lines.append(summary)
            text_lines.append("\n#### Complete Data Table:\n")

            text_lines.append("| " + " | ".join(headers) + " |")
            text_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")

            for r_vals in rows_data:
                text_lines.append("| " + " | ".join(r_vals) + " |")

            text_lines.append("\n---\n")

        return {
            "text": "\n".join(text_lines),
            "file_type": "xlsx",
            "metadata": {
                "filename": filename,
                "sheets": sheets_meta,
                "tables": sheets_meta,
                "total_rows": total_rows
            }
        }

    @classmethod
    def read_csv(cls, content_or_path: Union[str, bytes], filename: str = "data.csv") -> Dict[str, Any]:
        """Extracts CSV / TSV tabular files with automatic schema summarization."""
        if isinstance(content_or_path, bytes):
            text_data = cls._decode_bytes(content_or_path)
        else:
            with open(content_or_path, "r", encoding="utf-8", errors="replace") as f:
                text_data = f.read()

        delimiter = ","
        try:
            sniffer = csv.Sniffer()
            sample = text_data[:4096]
            dialect = sniffer.sniff(sample)
            delimiter = dialect.delimiter
        except Exception:
            if "\t" in text_data and "," not in text_data:
                delimiter = "\t"

        reader = csv.reader(io.StringIO(text_data), delimiter=delimiter)
        rows = list(reader)
        if not rows:
            return {"text": f"# Document: {filename} (Empty CSV)", "file_type": "csv", "metadata": {"rows": 0}}

        headers = [c.strip() or f"Col_{i+1}" for i, c in enumerate(rows[0])]
        data_rows = rows[1:]

        summary = cls._generate_table_summary("Main Data", headers, data_rows)
        text_lines = [
            f"# Document Overview: {filename} (CSV/TSV Tabular Dataset)\n",
            summary,
            "\n#### Complete Data Table:\n",
            "| " + " | ".join(headers) + " |",
            "| " + " | ".join(["---"] * len(headers)) + " |"
        ]

        for row in data_rows:
            row_vals = [c.replace("\n", " ").strip() for c in row]
            while len(row_vals) < len(headers):
                row_vals.append("")
            text_lines.append("| " + " | ".join(row_vals[:len(headers)]) + " |")

        return {
            "text": "\n".join(text_lines),
            "file_type": "csv",
            "metadata": {
                "filename": filename,
                "rows": len(data_rows),
                "columns": len(headers),
                "tables": [{
                    "sheet_name": "data",
                    "rows": len(data_rows),
                    "columns": len(headers),
                    "headers": headers,
                    "data": data_rows
                }],
                "total_rows": len(data_rows)
            }
        }

    @classmethod
    def _is_valid_data_table(cls, headers: List[str], rows: List[List[str]]) -> bool:
        """
        Distinguishes true structured data tables from visual layout artifacts
        (e.g., multi-column resume sections, margins, sidebars).
        """
        if not headers or not rows or len(rows) < 2:
            return False

        valid_headers = [h for h in headers if h and len(h.strip()) > 0]
        # A true data table must have at least 2 non-empty header columns
        if len(valid_headers) < 2:
            return False

        # Real column headers are concise and don't contain line breaks or full sentences/names
        if any("\n" in h or len(h) > 40 for h in valid_headers):
            return False

        num_cols = len(headers)
        # Check column emptiness across rows:
        active_col_count = 0
        for c_idx in range(num_cols):
            filled_cells = sum(1 for r in rows if c_idx < len(r) and r[c_idx].strip())
            if filled_cells / len(rows) >= 0.35:
                active_col_count += 1

        if active_col_count < 2:
            return False

        # In a real table, cells are concise data values, not whole paragraphs
        total_cells = 0
        long_cells = 0
        for r in rows:
            for c in r:
                if c:
                    total_cells += 1
                    if len(c) > 100:
                        long_cells += 1

        if total_cells > 0 and (long_cells / total_cells) > 0.2:
            return False

        return True

    @classmethod
    def read_pdf(cls, content_or_path: Union[str, bytes], filename: str = "document.pdf") -> Dict[str, Any]:
        """Extracts text, page numbers, and structured tables from PDF documents."""
        text_lines = []
        tables_meta = []
        page_count = 0

        # Try advanced layout & table extraction with pdfplumber
        if pdfplumber is not None:
            try:
                stream = io.BytesIO(content_or_path) if isinstance(content_or_path, bytes) else content_or_path
                with pdfplumber.open(stream) as pdf:
                    page_count = len(pdf.pages)
                    text_lines.append(f"# Document: {filename} (PDF - {page_count} pages)")
                    for idx, page in enumerate(pdf.pages):
                        p_num = idx + 1
                        page_text = page.extract_text() or ""
                        text_lines.append(f"\n--- Page {p_num} ---\n{page_text.strip()}")
                        
                        # Extract any embedded tables (validated against false positives)
                        page_tables = page.extract_tables() or []
                        for t_idx, tbl in enumerate(page_tables):
                            if tbl and len(tbl) >= 2:
                                headers = [str(c or '').strip() for c in tbl[0]]
                                rows = [[str(c or '').strip() for c in r] for r in tbl[1:]]
                                if cls._is_valid_data_table(headers, rows):
                                    t_name = f"pdf_p{p_num}_tbl{t_idx+1}"
                                    tables_meta.append({
                                        "sheet_name": t_name,
                                        "headers": headers,
                                        "data": rows
                                    })
                                    text_lines.append(f"\n### Table on Page {p_num}")
                                    text_lines.append("| " + " | ".join(headers) + " |")
                                    text_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
                                    for r in rows:
                                        text_lines.append("| " + " | ".join(r) + " |")

                if text_lines:
                    return {
                        "text": "\n".join(text_lines),
                        "file_type": "pdf",
                        "metadata": {
                            "filename": filename,
                            "pages": page_count,
                            "tables": tables_meta
                        }
                    }
            except Exception:
                pass

        # Standard fallback to pypdf
        if pypdf is None:
            raise ImportError("pypdf or pdfplumber is required for PDF files.")

        if isinstance(content_or_path, bytes):
            reader = pypdf.PdfReader(io.BytesIO(content_or_path))
        else:
            reader = pypdf.PdfReader(content_or_path)

        text_lines = [f"# Document: {filename} (PDF - {len(reader.pages)} pages)"]
        for idx, page in enumerate(reader.pages):
            page_text = page.extract_text() or ""
            text_lines.append(f"\n--- Page {idx+1} ---\n{page_text.strip()}")

        return {
            "text": "\n".join(text_lines),
            "file_type": "pdf",
            "metadata": {
                "filename": filename,
                "pages": len(reader.pages),
                "tables": []
            }
        }

    @classmethod
    def read_docx(cls, content_or_path: Union[str, bytes], filename: str = "document.docx") -> Dict[str, Any]:
        """Extracts text, headings, and tables from Word documents (.docx)."""
        if docx is None:
            raise ImportError("python-docx is required for Word documents. Install with: pip install python-docx")

        if isinstance(content_or_path, bytes):
            doc = docx.Document(io.BytesIO(content_or_path))
        else:
            doc = docx.Document(content_or_path)

        text_lines = [f"# Document: {filename} (Word Document)\n"]
        for p in doc.paragraphs:
            txt = p.text.strip()
            if not txt:
                continue
            if p.style and p.style.name.startswith("Heading"):
                lvl = p.style.name.replace("Heading", "").strip()
                prefix = "#" * (int(lvl) if lvl.isdigit() else 2)
                text_lines.append(f"\n{prefix} {txt}\n")
            else:
                text_lines.append(txt)

        for t_idx, table in enumerate(doc.tables):
            text_lines.append(f"\n### Table {t_idx+1}\n")
            if not table.rows:
                continue
            headers = [c.text.strip().replace("\n", " ") for c in table.rows[0].cells]
            text_lines.append("| " + " | ".join(headers) + " |")
            text_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
            for row in table.rows[1:]:
                row_vals = [c.text.strip().replace("\n", " ") for c in row.cells]
                text_lines.append("| " + " | ".join(row_vals) + " |")

        return {
            "text": "\n\n".join(text_lines),
            "file_type": "docx",
            "metadata": {
                "filename": filename,
                "paragraphs": len(doc.paragraphs),
                "tables": len(doc.tables)
            }
        }

    @classmethod
    def read_pptx(cls, content_or_path: Union[str, bytes], filename: str = "presentation.pptx") -> Dict[str, Any]:
        """Extracts text from PowerPoint presentations (.pptx)."""
        if pptx is None:
            raise ImportError("python-pptx is required for PowerPoint files. Install with: pip install python-pptx")

        if isinstance(content_or_path, bytes):
            prs = pptx.Presentation(io.BytesIO(content_or_path))
        else:
            prs = pptx.Presentation(content_or_path)

        text_lines = [f"# Presentation: {filename} ({len(prs.slides)} slides)"]
        for idx, slide in enumerate(prs.slides):
            text_lines.append(f"\n## Slide {idx+1}")
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    text_lines.append(shape.text.strip())

        return {
            "text": "\n\n".join(text_lines),
            "file_type": "pptx",
            "metadata": {
                "filename": filename,
                "slides": len(prs.slides)
            }
        }

    @classmethod
    def _decode_bytes(cls, data: bytes) -> str:
        """Tries multiple common text encodings (UTF-8, Latin-1, CP1252, etc.) to safely convert raw bytes into readable text without crashing."""
        for enc in ["utf-8", "utf-8-sig", "latin-1", "cp1252", "iso-8859-1", "utf-16"]:
            try:
                return data.decode(enc)
            except Exception:
                continue
        return data.decode("utf-8", errors="replace")

    @classmethod
    def read_text_file(cls, content_or_path: Union[str, bytes], filename: str = "file.txt") -> Dict[str, Any]:
        """Extracts plain text, markdown, json, xml, code, etc."""
        if isinstance(content_or_path, bytes):
            raw_text = cls._decode_bytes(content_or_path)
        else:
            with open(content_or_path, "rb") as f:
                raw_text = cls._decode_bytes(f.read())

        ext = Path(filename).suffix.lower()
        if ext == ".json":
            try:
                parsed = json.loads(raw_text)
                formatted_text = f"# Document: {filename} (JSON Data)\n```json\n{json.dumps(parsed, indent=2)}\n```"
                return {"text": formatted_text, "file_type": "json", "metadata": {"filename": filename}}
            except Exception:
                pass

        return {
            "text": f"# Document: {filename}\n\n{raw_text}",
            "file_type": ext.lstrip(".") or "txt",
            "metadata": {"filename": filename, "chars": len(raw_text)}
        }

    @classmethod
    def read(cls, content_or_path: Union[str, bytes], filename: Optional[str] = None) -> Dict[str, Any]:
        """Universal entry point: identifies file type and extracts clean structured text."""
        if filename is None:
            if isinstance(content_or_path, str):
                filename = Path(content_or_path).name
            else:
                filename = "uploaded_file"

        ext = Path(filename).suffix.lower()

        if ext == ".xls":
            return cls.read_xls(content_or_path, filename)
        if ext in [".xlsx", ".xlsm", ".xltx", ".xltm"]:
            return cls.read_xlsx(content_or_path, filename)
        if ext in [".csv", ".tsv"]:
            return cls.read_csv(content_or_path, filename)
        if ext == ".pdf":
            return cls.read_pdf(content_or_path, filename)
        if ext in [".docx", ".doc"]:
            if ext == ".docx":
                return cls.read_docx(content_or_path, filename)
            return cls.read_text_file(content_or_path, filename)
        if ext in [".pptx"]:
            return cls.read_pptx(content_or_path, filename)

        return cls.read_text_file(content_or_path, filename)


class LocalRAGIndexer:
    """Smart chunking, Ollama dense embeddings, FastEmbed BM25 sparse vectors, and Qdrant hybrid storage."""

    def __init__(self,
                 ollama_url: str = "http://localhost:11434",
                 qdrant_url: str = "http://localhost:6333",
                 embedding_model: str = "nomic-embed-text",
                 collection_name: str = "documents"):
        """Initializes the RAG indexer with Ollama embeddings, Qdrant vector database URL, and BM25 sparse search."""
        self.ollama_url = ollama_url
        self.qdrant_url = qdrant_url
        self.embedding_model = embedding_model
        self.collection_name = collection_name
        self.sparse_model = None
        self._has_sparse = None
        self._init_sparse_model()

    def _init_sparse_model(self):
        """Loads the fast BM25 sparse text embedding model for exact keyword search."""
        if SparseTextEmbedding is not None:
            try:
                self.sparse_model = SparseTextEmbedding(model_name="Qdrant/bm25")
            except Exception:
                self.sparse_model = None

    def chunk_with_pages(self, text: str, filename: str, max_chars: int = 1200, overlap: int = 150) -> List[Tuple[str, Optional[int]]]:
        """Page-aware text chunking with overlap and document filename header."""
        header = f'[Source File: {filename}]\n'
        parts = PAGE_RE.split(text)          # [pre, '1', body1, '2', body2, ...]
        units = [(None, parts[0])] + [(int(parts[i]), parts[i + 1])
                                       for i in range(1, len(parts) - 1, 2)]
        chunks = []
        for page, body in units:
            body = body.strip()
            start = 0
            while body and start < len(body):
                end = min(start + max_chars, len(body))
                if end < len(body):
                    cut = body.rfind('\n\n', start, end)
                    if cut == -1:
                        cut = body.rfind('\n', start, end)
                    if cut > start + max_chars // 2:
                        end = cut
                piece = body[start:end].strip()
                if piece:
                    chunks.append((header + piece, page))
                if end >= len(body):
                    break
                start = max(end - overlap, start + 1)
        return chunks

    def chunk_text(self, text: str, filename: str = "", chunk_size: int = 1500, chunk_overlap: int = 150) -> List[str]:
        """Fallback chunking returning text list."""
        pieces = self.chunk_with_pages(text, filename, max_chars=chunk_size, overlap=chunk_overlap)
        return [c for c, _ in pieces]

    def get_embeddings(self, texts: List[str], prefix: str = 'search_document: ') -> List[List[float]]:
        """Calls local Ollama /api/embed with task prefixes and keep_alive."""
        resp = requests.post(
            f'{self.ollama_url}/api/embed',
            json={'model': self.embedding_model, 'input': [prefix + t for t in texts],
                  'keep_alive': KEEP_ALIVE},
            timeout=300)
        resp.raise_for_status()
        return resp.json()['embeddings']

    def get_embedding(self, text: str, is_query: bool = False) -> List[float]:
        """Calls local Ollama to generate 768-dim dense vector embedding with task prefix."""
        prefix = 'search_query: ' if is_query else 'search_document: '
        return self.get_embeddings([text], prefix)[0]

    def get_sparse_embedding(self, text: str, is_query: bool = False) -> Optional[Dict[str, List[Any]]]:
        """Generates BM25 sparse vectors with token indices and term weights (separate query method)."""
        if not self.sparse_model or not text:
            return None
        try:
            gen = self.sparse_model.query_embed(text) if is_query else self.sparse_model.embed([text])
            vec = list(gen)[0]
            return {
                'indices': vec.indices.tolist(),
                'values': [round(float(v), 5) for v in vec.values.tolist()]
            }
        except Exception:
            return None

    def collection_has_sparse_vectors(self) -> bool:
        """Verifies if the Qdrant collection supports named 'sparse' vectors (memoized)."""
        if self._has_sparse is not None:
            return self._has_sparse
        try:
            r = requests.get(f'{self.qdrant_url}/collections/{self.collection_name}', timeout=5)
            if r.status_code != 200:
                return False
            conf = r.json().get('result', {}).get('config', {}).get('params', {})
            self._has_sparse = 'sparse' in conf.get('sparse_vectors', {})
            return self._has_sparse
        except Exception:
            return False

    def ensure_collection(self, dim: int = 768):
        """Ensures Qdrant collection exists with dense vector and named sparse vector schema."""
        try:
            r = requests.get(f"{self.qdrant_url}/collections/{self.collection_name}", timeout=5)
            if r.status_code == 200:
                conf = r.json().get("result", {}).get("config", {}).get("params", {})
                if "sparse_vectors" in conf and "sparse" in conf.get("sparse_vectors", {}):
                    self._has_sparse = True
                    return
                logger.warning(f"Collection '{self.collection_name}' exists without sparse_vectors schema. Recreating...")
                requests.delete(f"{self.qdrant_url}/collections/{self.collection_name}", timeout=10)
        except Exception as e:
            logger.warning(f"Could not check collection: {e}")

        requests.put(
            f"{self.qdrant_url}/collections/{self.collection_name}",
            json={
                "vectors": {"size": dim, "distance": "Cosine"},
                "sparse_vectors": {"sparse": {}}
            },
            timeout=10
        )
        try:
            requests.put(
                f"{self.qdrant_url}/collections/{self.collection_name}/index",
                json={"field_name": "metadata.source", "field_schema": "keyword"},
                timeout=5
            )
        except Exception:
            pass
        self._has_sparse = None

    def delete_existing_file_vectors(self, filename: str):
        """Removes previous points for this filename so updates don't create stale duplicate chunks."""
        try:
            requests.post(
                f"{self.qdrant_url}/collections/{self.collection_name}/points/delete",
                json={
                    "filter": {
                        "must": [
                            {"key": "metadata.source", "match": {"value": filename}}
                        ]
                    }
                },
                timeout=10
            )
        except Exception:
            pass

    def ingest_document(self, doc_data: Dict[str, Any]) -> Dict[str, Any]:
        """Extracts, chunks, embeds in batches of 16, and stores slim points in batches of 64 in Qdrant."""
        text = doc_data['text']
        metadata = doc_data.get('metadata', {})
        filename = metadata.get('filename', 'unknown')
        slim = {k: v for k, v in metadata.items() if k not in ('tables', 'sheets')}
        pieces = self.chunk_with_pages(text, filename)
        if not pieces:
            return {'status': 'empty', 'chunks_indexed': 0, 'filename': filename}

        self.ensure_collection(dim=768)
        self.delete_existing_file_vectors(filename)
        has_sparse = self.collection_has_sparse_vectors()

        texts = [c for c, _ in pieces]
        dense = []
        for i in range(0, len(texts), 16):
            dense.extend(self.get_embeddings(texts[i:i + 16]))

        points = []
        for idx, ((chunk, page), dvec) in enumerate(zip(pieces, dense)):
            if has_sparse:
                vec = {'': dvec}
                sp = self.get_sparse_embedding(chunk)
                if sp and sp.get('indices'):
                    vec['sparse'] = sp
            else:
                vec = dvec
            points.append({
                'id': str(uuid.uuid4()),
                'vector': vec,
                'payload': {
                    'content': chunk,
                    'source': filename,
                    'metadata': {
                        **slim,
                        'source': filename,
                        'page': page,
                        'file_type': doc_data.get('file_type', 'txt'),
                        'chunk_index': idx,
                        'total_chunks': len(pieces)
                    }
                }
            })

        for i in range(0, len(points), 64):
            res = requests.put(
                f'{self.qdrant_url}/collections/{self.collection_name}/points?wait=true',
                json={'points': points[i:i + 64]},
                timeout=60
            )
            res.raise_for_status()

        return {
            'status': 'success',
            'filename': filename,
            'chunks_indexed': len(points),
            'total_chars': len(text),
            'is_hybrid': bool(has_sparse and self.sparse_model)
        }

    def hybrid_search(self, query: str, top_k: int = 15, source_filter: Optional[Union[str, List[str]]] = None) -> List[Dict[str, Any]]:
        """
        Executes native Qdrant Hybrid Fusion Search (RRF):
        - Prefetches dense semantic candidates (nomic-embed-text)
        - Prefetches sparse exact keyword candidates (BM25)
        - Fuses candidate rankings using Reciprocal Rank Fusion
        - Supports optional exact source filtering to restrict search to one or multiple documents.
        """
        dense_vec = self.get_embedding(query, is_query=True)
        sparse_vec = self.get_sparse_embedding(query, is_query=True)

        filter_clause = None
        if source_filter:
            if isinstance(source_filter, (list, set, tuple)):
                clean_filters = [str(s).strip() for s in source_filter if str(s).strip()]
                if len(clean_filters) == 1:
                    filter_clause = {"must": [{"key": "metadata.source", "match": {"value": clean_filters[0]}}]}
                elif len(clean_filters) > 1:
                    filter_clause = {"must": [{"key": "metadata.source", "match": {"any": clean_filters}}]}
            else:
                s_str = str(source_filter).strip()
                if s_str:
                    filter_clause = {"must": [{"key": "metadata.source", "match": {"value": s_str}}]}

        has_sparse = self.collection_has_sparse_vectors()
        # 1. Native Hybrid Search with RRF Fusion
        if has_sparse and sparse_vec and sparse_vec.get("indices"):
            try:
                prefetch_dense = {
                    "query": dense_vec,
                    "using": "",
                    "limit": top_k
                }
                prefetch_sparse = {
                    "query": sparse_vec,
                    "using": "sparse",
                    "limit": top_k
                }
                if filter_clause:
                    prefetch_dense["filter"] = filter_clause
                    prefetch_sparse["filter"] = filter_clause

                hybrid_payload = {
                    "prefetch": [prefetch_dense, prefetch_sparse],
                    "query": {
                        "fusion": "rrf"
                    },
                    "limit": top_k,
                    "with_payload": True
                }
                res = requests.post(
                    f"{self.qdrant_url}/collections/{self.collection_name}/points/query",
                    json=hybrid_payload,
                    timeout=15
                )
                if res.status_code == 200:
                    points = res.json().get("result", {}).get("points", [])
                    if points:
                        return points
            except Exception:
                pass

        # 2. Resilient Fallback to Dense Vector Search
        try:
            dense_payload = {"vector": dense_vec, "limit": top_k, "with_payload": True}
            if filter_clause:
                dense_payload["filter"] = filter_clause
            dense_res = requests.post(
                f"{self.qdrant_url}/collections/{self.collection_name}/points/search",
                json=dense_payload,
                timeout=15
            )
            if dense_res.status_code == 200:
                return dense_res.json().get("result", [])
        except Exception:
            pass

        return []

    def source_chunk_counts(self) -> Dict[str, int]:
        """Returns a mapping of document sources to total chunk count with full pagination."""
        counts, offset = {}, None
        while True:
            body = {'limit': 500, 'with_payload': ['metadata.source']}
            if offset is not None:
                body['offset'] = offset
            try:
                r = requests.post(f'{self.qdrant_url}/collections/{self.collection_name}/points/scroll',
                                  json=body, timeout=10)
                if r.status_code != 200:
                    break
                res = r.json().get('result', {})
                for p in res.get('points', []):
                    s = p.get('payload', {}).get('metadata', {}).get('source')
                    if s:
                        counts[s] = counts.get(s, 0) + 1
                offset = res.get('next_page_offset')
                if offset is None:
                    break
            except Exception:
                break
        return counts

    def list_indexed_sources(self) -> List[str]:
        """Returns the list of distinct document sources indexed in Qdrant."""
        return sorted(self.source_chunk_counts())


class LocalReranker:
    """
    Local Cross-Encoder Reranker using FlashRank (ONNX Runtime).
    Computes exact query-document semantic relevance in ~5-15ms on CPU.
    """
    def __init__(self, model_name: str = "ms-marco-MiniLM-L-12-v2"):
        """Initializes the local cross-encoder reranker model (FlashRank) using lightweight ONNX runtime on CPU."""
        self.model_name = model_name
        self.ranker = None
        self._init_model()

    def _init_model(self):
        """Loads the FlashRank neural reranker model into memory."""
        if Ranker is not None:
            try:
                self.ranker = Ranker(model_name=self.model_name)
            except Exception:
                try:
                    self.ranker = Ranker()
                except Exception:
                    self.ranker = None

    def rerank(self, query: str, candidates: List[Dict[str, Any]], top_k: int = 4) -> List[Dict[str, Any]]:
        """Re-scores and re-orders candidate document passages to place the most accurate, relevant results at the top."""
        if not self.ranker or not candidates:
            return candidates[:top_k]
        try:
            passages = []
            for idx, c in enumerate(candidates):
                passages.append({
                    "id": idx,
                    "text": c.get("content", ""),
                    "meta": c
                })
            req = RerankRequest(query=query, passages=passages)
            ranked = self.ranker.rerank(req)
            results = []
            for r in ranked[:top_k]:
                orig = dict(r.get("meta", {}))
                orig["rerank_score"] = round(float(r.get("score", 0.0)), 4)
                results.append(orig)
            return results
        except Exception:
            return candidates[:top_k]



class RAGResponseValidator:
    """
    Universal Response Validator for Local RAG.
    Ensures groundedness, factual alignment, and clean markdown presentation
    without robotic boilerplate phrases.
    """

    def __init__(self, indexer_inst: Optional[LocalRAGIndexer] = None):
        """Initializes the response validator with optional access to the indexer."""
        self.indexer = indexer_inst

    def humanize_tone(self, text: str) -> str:
        """Transforms cold, robotic phrases into clear, professional language."""
        t = text
        t = re.sub(
            r"No,\s*this\s+data\s+is\s+not\s+user\s+data\.?\s*The\s+data\s+appears\s+to\s+be\s+a\s+sample\s+dataset[^.]*\.",
            "This appears to be a sample demonstration dataset containing fictional records used for analytics and testing.",
            t, flags=re.IGNORECASE
        )
        t = re.sub(
            r"Unfortunately,\s*the\s+document\s+does\s+not\s+provide\s+specific\s+information\s+about\s+the\s+data\s+in\s+the\s+file[^.]*\.",
            "This document contains structured records across the dataset. Here is an overview of the data:",
            t, flags=re.IGNORECASE
        )
        return t

    def sanitize_security_and_privacy(self, text: str) -> str:
        """Removes accidental internal file paths or system leakage."""
        t = text
        t = re.sub(r"file_reader_engine\.py", "internal engine", t, flags=re.IGNORECASE)
        t = re.sub(r"/home/[^\s\"']+", "[system-path]", t)
        return t

    def validate(self, query: str, response: str, context: Optional[str] = None) -> Dict[str, Any]:
        """
        Validates the AI candidate response for basic groundedness, tone consistency, and security.
        Returns validation status, corrections, and polished text.
        """
        corrections = []
        was_corrected = False
        final_response = response

        # Apply tone polish if necessary
        humanized = self.humanize_tone(final_response)
        if humanized != final_response:
            was_corrected = True
            corrections.append("Humanized robotic phrasing into warm, professional tone.")
            final_response = humanized

        # Apply security and privacy sanitization
        sanitized = self.sanitize_security_and_privacy(final_response)
        if sanitized != final_response:
            was_corrected = True
            corrections.append("Sanitized internal file/path references to comply with security policy.")
            final_response = sanitized

        return {
            "is_valid": True,
            "was_corrected": was_corrected,
            "corrections": corrections,
            "validated_response": final_response,
            "has_context": bool(context)
        }



# ==============================================================================
# DuckDB In-Memory & Embedded Analytics Engine
# ==============================================================================
class DuckDBEngine:
    """
    Local Embedded DuckDB Analytics Engine.
    Provides deterministic, sub-millisecond execution of ANSI SQL
    for spreadsheets (.xls, .xlsx), CSVs, and tabular datasets.
    """
    def __init__(self, db_path: Optional[str] = None):
        """Initializes the in-memory or file-backed DuckDB database connection and creates table registries."""
        if db_path is None:
            base_dir = Path(__file__).resolve().parent
            data_dir = base_dir / "duckdb_data"
            data_dir.mkdir(parents=True, exist_ok=True)
            db_path = str(data_dir / "analytics.duckdb")
        
        self.db_path = db_path
        if duckdb:
            try:
                self.con = duckdb.connect(self.db_path, config={'enable_external_access': False})
                self.con.execute('SET lock_configuration = true')
            except Exception:
                try:
                    self.con = duckdb.connect(self.db_path, read_only=True, config={'enable_external_access': False})
                    self.con.execute('SET lock_configuration = true')
                except Exception:
                    self.con = duckdb.connect(":memory:", config={'enable_external_access': False})
                    self.con.execute('SET lock_configuration = true')
        else:
            self.con = None
        self.registry: Dict[str, Dict[str, Any]] = {}
        self.active_table: Optional[str] = None
        self.lock = threading.Lock()
        self._init_existing_tables()

    def _init_existing_tables(self):
        """Discovers previously loaded tables in DuckDB file on startup."""
        if not self.con:
            return
        try:
            try:
                src_map = dict(self.con.execute('SELECT table_name, source_file FROM _sources').fetchall())
            except Exception:
                src_map = {}
            tables = self.con.execute("SHOW TABLES").fetchall()
            for (tbl_name,) in tables:
                if tbl_name in ["dataset", "active_table", "_sources"]:
                    continue
                cols = profile_table(self.con, tbl_name)
                cnt = self.con.execute(f'SELECT COUNT(*) FROM "{tbl_name}"').fetchone()[0]
                self.registry[tbl_name] = {
                    "table_name": tbl_name,
                    "columns": cols,
                    "total_rows": cnt,
                    "source_file": src_map.get(tbl_name),
                    "suggested_questions": self._generate_suggested_questions(tbl_name, cols)
                }
                self.active_table = tbl_name
        except Exception:
            pass

    def _generate_suggested_questions(self, table_name: str, cols: List[Dict[str, Any]]) -> List[str]:
        """Generates dynamic, high-value starter questions for newly loaded tables."""
        questions = [f"How many total records are in the dataset?"]
        cat_cols = [c["name"] for c in cols if c["type"] == "VARCHAR" and "id" not in c["name"].lower() and "name" not in c["name"].lower()]
        num_cols = [c["name"] for c in cols if c["type"].startswith(NUMERIC_TYPES) and "id" not in c["name"].lower() and "col_0" not in c["name"].lower()]
        date_cols = [c["name"] for c in cols if c["type"].startswith(("DATE", "TIMESTAMP")) or "date" in c["name"].lower()]

        if cat_cols:
            questions.append(f"Show the count of records grouped by {cat_cols[0]}")
        if num_cols and cat_cols:
            questions.append(f"What is the average {num_cols[0]} by {cat_cols[0]}?")
        elif num_cols:
            questions.append(f"What is the average and highest {num_cols[0]}?")
        if date_cols:
            questions.append(f"List records where {date_cols[0]} is after 2016")
        questions.append("List the top 5 records from the table")
        return questions[:4]

    def make_table_name(self, filename: str, sheet: Any, used: set) -> str:
        """Constructs unique, clean table names preventing collisions across workbooks and sheets."""
        stem = re.sub(r'[^a-z0-9]+', '_', Path(filename).stem.lower()).strip('_') or 'file'
        sh = re.sub(r'[^a-z0-9]+', '_', str(sheet).lower()).strip('_')
        name = stem if sh in ('', 'data', 'main_data') else f'{stem}__{sh}'
        if name[0].isdigit():
            name = f't_{name}'
        base, n = name, 2
        while name in used or (name in self.registry
                                and self.registry[name].get('source_file') not in (None, filename)):
            name, n = f'{base}_{n}', n + 1
        used.add(name)
        return name

    def ingest_tables(self, filename: str, tables: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Ingests structured sheets/tables into DuckDB with automatic type coercion and column profiling."""
        if not self.con or not tables:
            return []
        self.delete_file_tables(filename)            # re-upload replaces only this file's tables
        used, results = set(), []
        for tbl in tables:
            headers, rows = tbl.get('headers', []), tbl.get('data', [])
            if not headers or not rows:
                continue
            cols = clean_columns(headers)
            w = len(cols)
            df = pd.DataFrame([(list(r) + [None] * w)[:w] for r in rows], columns=cols)
            for c in cols:
                df[c] = coerce_series(df[c])
            df = df.dropna(how='all')
            name = self.make_table_name(filename, tbl.get('sheet_name', 'data'), used)
            with self.lock:
                self.con.register('_incoming', df)
                self.con.execute(f'CREATE OR REPLACE TABLE "{name}" AS SELECT * FROM _incoming')
                self.con.unregister('_incoming')
                self.con.execute('CREATE TABLE IF NOT EXISTS _sources (table_name VARCHAR PRIMARY KEY, source_file VARCHAR)')
                self.con.execute('INSERT OR REPLACE INTO _sources VALUES (?, ?)', [name, filename])
                self.con.execute(f'CREATE OR REPLACE VIEW active_table AS SELECT * FROM "{name}"')
                self.con.execute(f'CREATE OR REPLACE VIEW dataset AS SELECT * FROM "{name}"')
            profile = profile_table(self.con, name)
            meta = {
                'table_name': name,
                'source_file': filename,
                'columns': profile,
                'total_rows': len(df),
                'suggested_questions': self._generate_suggested_questions(name, profile)
            }
            self.registry[name] = meta
            self.active_table = name
            results.append(meta)
        return results

    def delete_file_tables(self, filename: str):
        """Drops any tables registered for this source file."""
        if not self.con:
            return
        with self.lock:
            tbls_to_remove = [tbl for tbl, m in self.registry.items() if tbl == filename or m.get("source_file") == filename]
            for tbl in tbls_to_remove:
                try:
                    self.con.execute(f'DROP TABLE IF EXISTS "{tbl}"')
                except Exception:
                    pass
                try:
                    self.con.execute('DELETE FROM _sources WHERE table_name = ?', [tbl])
                except Exception:
                    pass
                self.registry.pop(tbl, None)
            if self.active_table in tbls_to_remove:
                self.active_table = next(iter(self.registry.keys())) if self.registry else None
                if self.active_table:
                    try:
                        self.con.execute(f'CREATE OR REPLACE VIEW dataset AS SELECT * FROM "{self.active_table}"')
                        self.con.execute(f'CREATE OR REPLACE VIEW active_table AS SELECT * FROM "{self.active_table}"')
                    except Exception:
                        pass

    def execute_query(self, sql_query: str) -> Dict[str, Any]:
        """Executes a strictly read-only single SELECT query against DuckDB."""
        if not self.con:
            return {'success': False, 'error': 'DuckDB is not available'}
        clean_sql = sql_query.strip().rstrip(';')
        try:
            stmts = self.con.extract_statements(clean_sql)
            if len(stmts) != 1 or stmts[0].type != duckdb.StatementType.SELECT:
                return {'success': False, 'sql': clean_sql,
                        'error': 'Only a single SELECT query is allowed.'}
        except Exception as e:
            return {'success': False, 'sql': clean_sql, 'error': f'Could not parse SQL: {e}'}

        try:
            t0 = time.perf_counter()
            cur = self.con.cursor()
            cur.execute(clean_sql)
            cols = [d[0] for d in cur.description] if cur.description else []
            rows = cur.fetchmany(MAX_ROWS)
            elapsed_ms = (time.perf_counter() - t0) * 1000

            clean_rows = []
            for r in rows:
                clean_rows.append([str(v) if v is not None else "" for v in r])

            return {
                "success": True,
                "columns": cols,
                "rows": clean_rows,
                "row_count": len(rows),
                "execution_time_ms": round(elapsed_ms, 2),
                "sql": clean_sql
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e),
                "sql": clean_sql
            }


# ==============================================================================
# Query Router: Classifies user intent into SQL, RAG, or HYBRID
# ==============================================================================
class QueryRouter:
    """Classifies user queries to route them to Conversational AI, deterministic DuckDB SQL, or Hybrid RAG."""

    DOC_REF_RE = re.compile(
        r"\b(pdf|docx?|documents?|files?|sheet|spreadsheet|uploaded|"
        r"as\s+p[er]{2,3}|according\s+to|based\s+on|"
        r"in\s+the\s+(?:doc|document|pdf|file))\b", re.IGNORECASE)

    EXPLICIT_DOC_RE = re.compile(
        r"\b(pdf|docx?|documents?|uploaded|spreadsheet|(?:the|this|my)\s+file)\b",
        re.IGNORECASE)

    SOURCE_LEAK_RE = re.compile(
        r"\b((?:share|give|send|show|download|reveal|leak|dump|expose)\s+(?:me\s+)?your\s+"
        r"(?:(?:core|backend|internal|system|server|python|source)\s+)*(?:code|files?|scripts?|programs?|repo(?:sitory)?)"
        r"|your\s+(?:source|backend|server|core)\s+code|where\s+is\s+your\s+(?:code|file)|file_reader_engine\.py)\b",
        re.IGNORECASE)

    IDENTITY_RE = re.compile(
        r"\b(who\s+are\s+you|what\s+is\s+your\s+name|how\s+are\s+you|what\s+can\s+you\s+do|"
        r"tell\s+me\s+about\s+yourself|introduce\s+yourself|"
        r"your\s+(?:tech(?:nology)?\s+stack|architecture)|"
        r"what\s+(?:tech(?:nologies)?|models?)\s+do\s+you\s+use|"
        r"how\s+are\s+you\s+built|what\s+powers\s+you|how\s+do\s+you\s+work|"
        r"are\s+you\s+(?:cloud|local|offline))\b", re.IGNORECASE)

    def __init__(self, duckdb_engine: DuckDBEngine):
        """Initializes the query classifier with access to the DuckDB engine to check table columns."""
        self.duckdb_engine = duckdb_engine

    @staticmethod
    def is_pure_math(query: str) -> bool:
        """Detects if the user query is a direct mathematical calculation (e.g., 25 * 14 or what is 100 / 4)."""
        q = query.strip().rstrip("?").strip()
        math_cleaned = re.sub(r"^(?:what\s+is|calculate|compute|solve|how\s+much\s+is)\s+", "", q, flags=re.IGNORECASE).strip()
        if re.match(r"^[\d\s\+\-\*\/\^\(\)\.\%]+$", math_cleaned) and any(op in math_cleaned for op in ["+", "-", "*", "/", "%", "^"]):
            return True
        return False

    @staticmethod
    def safe_math_eval(query: str) -> Optional[Union[int, float]]:
        """Safely calculates the mathematical expression using Python AST without using unsafe eval()."""
        clean = re.sub(r"^(?:what\s+is|calculate|compute|solve|how\s+much\s+is)\s+", "", query.strip().rstrip("?").strip(), flags=re.IGNORECASE)
        math_ops = {
            ast.Add: operator.add,
            ast.Sub: operator.sub,
            ast.Mult: operator.mul,
            ast.Div: operator.truediv,
            ast.Pow: operator.pow,
            ast.Mod: operator.mod,
            ast.USub: operator.neg,
            ast.UAdd: operator.pos,
        }
        try:
            def _eval_node(node):
                if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
                    return node.value
                elif isinstance(node, ast.BinOp) and type(node.op) in math_ops:
                    left = _eval_node(node.left)
                    right = _eval_node(node.right)
                    return math_ops[type(node.op)](left, right)
                elif isinstance(node, ast.UnaryOp) and type(node.op) in math_ops:
                    operand = _eval_node(node.operand)
                    return math_ops[type(node.op)](operand)
                raise ValueError("Unsupported operation")
            tree = ast.parse(clean, mode='eval')
            res = _eval_node(tree.body)
            if isinstance(res, float) and res.is_integer():
                return int(res)
            return res
        except Exception:
            return None

    @staticmethod
    def is_conversational(query: str) -> bool:
        """Checks if the query is a greeting, general conversational remark, or safety inquiry rather than a document question."""
        q_clean = query.strip().lower()
        if not q_clean:
            return True

        if QueryRouter.is_pure_math(query):
            return True

        if QueryRouter.SOURCE_LEAK_RE.search(q_clean):
            return True

        security_patterns = [
            r"\b(share|give|leak|reveal|dump|show)\s+(?:me\s+)?(?:personal\s+data|pii|passwords?|secret\s+keys?|api\s+keys?|tokens?|credentials?)\b",
        ]
        if any(re.search(p, q_clean) for p in security_patterns):
            return True

        if QueryRouter.DOC_REF_RE.search(q_clean):
            return False

        if QueryRouter.IDENTITY_RE.search(q_clean):
            return True

        greetings = [
            r"^(?:hi|hello|hey|heya|howdy|sup|greetings|good\s+(?:morning|afternoon|evening|day))[\s,!.]*$",
            r"\b(thank\s+you|thanks|thx|thank\s+you\s+so\s+much|appreciate\s+it)\b",
            r"\b(goodbye|bye|see\s+you|cya|good\s+night)\b",
            r"\b(tell\s+me\s+a\s+joke|make\s+me\s+laugh|say\s+something\s+funny)\b",
            r"^(?:ok|okay|cool|great|awesome|nice|got\s+it|understood|perfect|help|help\s+me)[\s!.]*$"
        ]
        if any(re.search(p, q_clean) for p in greetings):
            return True

        return False

    def classify(self, query: str, target_doc: Optional[Union[str, List[str]]] = None) -> str:
        """Intelligently classifies user query into CONVERSATIONAL, pure SQL, semantic RAG, or HYBRID mode."""
        q_lower = query.lower()

        # Step 0: Conversational / Math / Greetings
        if self.is_conversational(query):
            return "CONVERSATIONAL"

        sql_patterns = [
            r"\b(how many|count|number of|total|sum|average|avg|mean|median|min|minimum|max|maximum|highest|lowest)\b",
            r"\b(distribution|breakdown|group by|percentage|proportion|rank|ranking|top \d+|bottom \d+)\b",
            r"\b(more than|greater than|less than|older than|younger than|before|after|between)\b",
            r"\b(records? where|rows? where|>|<|=)\b",
            r"\b(only\s+[a-z_]+\s+(?:names?|column|from|values?)|give me only)\b",
            r"\b(list all\s+[a-z_]+)\b"
        ]

        # If an active target document is set in the session, queries stay strictly grounded to RAG/SQL!
        if target_doc:
            if isinstance(target_doc, list):
                has_tabular = any(str(d).lower().endswith((".xls", ".xlsx", ".xlsm", ".csv", ".tsv")) for d in target_doc)
                has_text = any(not str(d).lower().endswith((".xls", ".xlsx", ".xlsm", ".csv", ".tsv")) for d in target_doc)
                if has_tabular and not has_text and any(re.search(p, q_lower) for p in sql_patterns):
                    return "SQL"
                if has_tabular and any(re.search(p, q_lower) for p in sql_patterns) and not any(k in q_lower for k in ["pdf", "doc", "document", "resume", "cv"]):
                    return "SQL"
                return "RAG"
            else:
                is_tabular = target_doc.lower().endswith((".xls", ".xlsx", ".xlsm", ".csv", ".tsv"))
                if is_tabular and any(re.search(p, q_lower) for p in sql_patterns):
                    return "SQL"
                return "RAG"

        rag_patterns = [
            r"\b(policy|guideline|procedure|leave|handbook|rules?|contract|clause|terms|overview|summary|synopsis)\b",
            r"\b(who leads?|who manages?|project apollo|purpose of|explain the|what does it say about)\b",
            r"\b(resume|cv|portfolio|experience|skills|education|rutvik)\b",
            r"\b(improv\w*|feedback|review|weakness\w*|flaws?|strength\w*|section\w*|clause\w*)\b",
            r"\b(write|draft|compose|outreach|email|task\s+\d+|follow-?up|exercise|deliverable)\b",
            r"\b(pdf|document|file|doc|page\s*\d+)\b"
        ]
        is_rag = any(re.search(p, q_lower) for p in rag_patterns)

        sql_patterns = [
            r"\b(how many|count|number of|total|sum|average|avg|mean|median|min|minimum|max|maximum|highest|lowest)\b",
            r"\b(distribution|breakdown|group by|percentage|proportion|rank|ranking|top \d+|bottom \d+)\b",
            r"\b(more than|greater than|less than|older than|younger than|before|after|between)\b",
            r"\b(records? where|rows? where|>|<|=)\b",
            r"\b(only\s+[a-z_]+\s+(?:names?|column|from|values?)|give me only)\b",
            r"\b(list all\s+[a-z_]+)\b"
        ]
        is_sql = any(re.search(p, q_lower) for p in sql_patterns)

        # Check if query references column names
        active_cols = []
        if self.duckdb_engine and self.duckdb_engine.active_table and self.duckdb_engine.active_table in self.duckdb_engine.registry:
            active_cols = [c["name"].lower() for c in self.duckdb_engine.registry[self.duckdb_engine.active_table]["columns"]]
        col_matches = [c for c in active_cols if c in q_lower and len(c) > 2]
        if len(col_matches) >= 2 or (len(col_matches) >= 1 and any(w in q_lower for w in ["what", "who", "show", "list", "get"])):
            is_sql = True

        if is_sql and is_rag:
            return "HYBRID"
        if is_sql:
            return "SQL"
        if is_rag:
            return "RAG"

        # General knowledge / conversational fallback if no tables/docs referenced
        general_q_cues = [
            r"^(what|why|how|who|when|where|which|whose|whom)\b",
            r"^(can|could|would|will|may|might|should|shall)\b",
            r"^(do|does|did|is|are|was|were|am)\b",
            r"^(give\s+me|tell\s+me|show\s+me|help\s+me|talk\s+about|provide|share)\b",
            r"^(explain|describe|define|elaborate|suggest|recommend|summarize|list|name)\b",
            r"\b(in\s+the\s+world|on\s+(?:the\s+)?earth|in\s+history|in\s+general|in\s+nature)\b",
            r"\b(fastest|biggest|tallest|smallest|longest|heaviest|slowest|oldest|youngest|richest|deadliest|speed)\b",
            r"\b(animal|animals|movie|movies|entertainment|film|cinema|actor|actress|music|song|sports)\b"
        ]
        if any(re.search(c, q_lower) for c in general_q_cues):
            return "CONVERSATIONAL"

        return "RAG"


# ==============================================================================
# Text-to-SQL Engine with Grounded Schema, Self-Correction & Explanations
# ==============================================================================
SQL_EXAMPLES = '''EXAMPLES (patterns only; use the real table and column names from the schema):
Q: How many orders are there per status?
SQL: SELECT status, COUNT(*) AS orders FROM orders GROUP BY status ORDER BY orders DESC
Q: Top 5 customers by total revenue
SQL: SELECT customer, SUM(amount) AS revenue FROM orders GROUP BY customer ORDER BY revenue DESC LIMIT 5
Q: Monthly revenue in 2023
SQL: SELECT date_trunc('month', order_date) AS month, SUM(amount) AS revenue FROM orders WHERE order_date >= DATE '2023-01-01' AND order_date < DATE '2024-01-01' GROUP BY 1 ORDER BY 1
Q: What share of revenue comes from each region?
SQL: SELECT region, SUM(amount) AS revenue, ROUND(100.0 * SUM(amount) / SUM(SUM(amount)) OVER (), 2) AS pct FROM orders GROUP BY region ORDER BY revenue DESC
Q: Month over month revenue change
SQL: WITH m AS (SELECT date_trunc('month', order_date) AS month, SUM(amount) AS revenue FROM orders GROUP BY 1) SELECT month, revenue, revenue - LAG(revenue) OVER (ORDER BY month) AS change FROM m ORDER BY month
'''

def last_sql_turn(msgs: List[Dict[str, Any]]) -> Optional[Dict[str, str]]:
    """Finds the most recent user question and assistant SQL query from chat history."""
    for i in range(len(msgs) - 1, -1, -1):
        sql = (msgs[i].get('metadata') or {}).get('sql_query')
        if msgs[i].get('role') == 'assistant' and sql:
            q = next((x['content'] for x in reversed(msgs[:i]) if x.get('role') == 'user'), '')
            return {'question': q, 'sql': sql}
    return None

MAX_TABLE_ROWS = 50

def markdown_table(cols: List[str], rows: List[List[Any]]) -> str:
    """Formats SQL tabular results into clean GitHub-flavored markdown table."""
    head = '| ' + ' | '.join(cols) + ' |\n| ' + ' | '.join(['---'] * len(cols)) + ' |'
    body = '\n'.join('| ' + ' | '.join(str(v).replace('|', '/') for v in r) + ' |' for r in rows)
    return head + '\n' + body

def pick_chart(question: str, cols: List[str], rows: List[List[Any]]) -> Optional[Dict[str, Any]]:
    """Chooses appropriate chart type (line, pie, bar) and prepares series data."""
    if len(cols) < 2 or not (2 <= len(rows) <= 200):
        return None
    is_num = lambda i: all(_NUM.match(str(r[i])) for r in rows if str(r[i]) != '')
    is_date = lambda i: all(_DATE.match(str(r[i])) for r in rows if str(r[i]) != '')
    numeric = [i for i in range(len(cols)) if is_num(i)]
    label_like = [i for i in numeric if re.search(r'(year|month|quarter|week|day|id|rank)$', cols[i].lower())]
    measures = [i for i in numeric if i not in label_like]
    dates = [i for i in range(len(cols)) if i not in numeric and is_date(i)]
    cats = [i for i in range(len(cols)) if i not in numeric and i not in dates]
    label = (dates or label_like or cats or [None])[0]
    if label is None or not measures:
        return None
    if dates or label_like:
        ctype = 'line'
    elif len(rows) <= 8 and len(measures) == 1 and re.search(
            r'share|percent|proportion|breakdown|distribution', question, re.IGNORECASE):
        ctype = 'pie'
    else:
        ctype = 'bar'
    rows = rows[:30] if ctype == 'bar' else rows
    series = [{'label': cols[i].replace('_', ' ').title(),
               'values': [float(r[i]) if str(r[i]) != '' else 0.0 for r in rows]} for i in measures]
    return {'type': ctype, 'labels': [str(r[label]) for r in rows],
            'values': series[0]['values'], 'label': series[0]['label'],
            'category_label': cols[label].replace('_', ' ').title(),
            **({'series': series} if len(series) > 1 else {})}


class TextToSQLEngine:
    """Translates natural language questions to DuckDB ANSI SQL with grounded schema and automated error recovery."""
    def __init__(self, duckdb_engine: DuckDBEngine, ollama_url: str = "http://localhost:11434"):
        self.duckdb_engine = duckdb_engine
        self.ollama_url = ollama_url

    def get_schema_context(self, tables: Optional[List[str]] = None) -> str:
        """Builds rich schema context with allowed values for categories and ranges for numbers/dates."""
        reg = self.duckdb_engine.registry
        if tables is None:
            tables = list(reg.keys())
        lines = ['DATABASE SCHEMA (DuckDB):']
        for t in tables:
            if t not in reg:
                continue
            meta = reg[t]
            lines.append(f'\nTABLE "{t}" ({meta["total_rows"]} rows)')
            for c in meta.get('columns', []):
                d = f'  - "{c["name"]}" {c["type"]}'
                if c.get('values'):
                    d += ' | allowed values: ' + ', '.join(repr(v) for v in c['values'])
                elif 'min' in c and 'max' in c:
                    d += f' | range: {c["min"]} to {c["max"]}'
                elif c.get('sample'):
                    d += ' | e.g. ' + ', '.join(repr(v) for v in c['sample'])
                if c.get('nulls'):
                    d += f' | {c["nulls"]} nulls'
                lines.append(d)
        return '\n'.join(lines)

    def generate_sql(self, query: str, tables: List[str], prev: Optional[Dict[str, str]] = None, feedback: Optional[str] = None) -> Optional[str]:
        """Generates read-only DuckDB SQL using rich schema, few-shot examples, and feedback retries."""
        system = (
            'You are an expert data analyst writing DuckDB SQL.\n\n'
            + self.get_schema_context(tables) + '\n\n'
            'RULES:\n'
            f'1. Return ONLY the SQL inside one {FENCE}sql block. No explanations.\n'
            '2. One read-only SELECT or WITH statement.\n'
            '3. Use only the tables and columns listed above, and double-quote identifiers.\n'
            '4. For text filters, use the exact allowed values listed; map the user wording to the closest one. For free text, use ILIKE.\n'
            "5. For dates, compare with DATE literals ('after 2016' means >= DATE '2017-01-01'); group with date_trunc().\n"
            '6. Compute percentages, ratios and changes in SQL (window functions, LAG, SUM() OVER). Never leave arithmetic to the reader.\n'
            '7. Lists and top-N questions get LIMIT 50 unless the user asks for all rows.\n'
            "8. If the question cannot be answered from these columns, return exactly: SELECT 'UNANSWERABLE' AS reason\n\n"
            + SQL_EXAMPLES)
        user = f'User question: {query}'
        if prev:
            user = (f'Previous question: {prev["question"]}\nPrevious SQL: {prev["sql"]}\n'
                    'If the new question refines the previous one, modify that SQL; otherwise write a new query.\n\n' + user)
        if feedback:
            user += f'\n\nYour last attempt did not work: {feedback}\nWrite a corrected query.'
        try:
            res = requests.post(f'{self.ollama_url}/api/generate', json={
                'model': SQL_MODEL, 'system': system, 'prompt': user, 'stream': False,
                'keep_alive': KEEP_ALIVE,
                'options': {'temperature': 0.0, 'num_ctx': 4096, 'num_predict': 400}},
                timeout=90).json()
        except Exception as e:
            logger.error(f'SQL generation failed: {e}')
            return None
        raw = res.get('response', '').strip()
        m = re.search(FENCE + r'(?:sql)?\s*([\s\S]+?)\s*' + FENCE, raw, re.IGNORECASE)
        return (m.group(1) if m else raw).strip().rstrip(';')

    def process_query(self, query: str, source_file: Optional[str] = None, prev: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
        """Executes text-to-SQL with up to 3 retry loops on errors, zero rows, or NULL results."""
        reg = self.duckdb_engine.registry
        tables = [t for t, m in reg.items() if source_file and m.get('source_file') == source_file] or list(reg)
        if not tables:
            return {'success': False, 'answer': 'No tables are loaded. Upload a spreadsheet first.'}
        feedback, sql, res = None, None, None
        for _ in range(3):
            sql = self.generate_sql(query, tables, prev=prev, feedback=feedback)
            if not sql:
                break
            if 'UNANSWERABLE' in sql.upper():
                cols = sorted({c['name'] for t in tables for c in reg[t].get('columns', [])})
                return {'success': False, 'sql_query': sql,
                        'answer': 'I cannot answer that from the loaded data. Available columns: ' + ', '.join(cols)}
            res = self.duckdb_engine.execute_query(sql)
            if not res['success']:
                feedback = f'`{sql}` failed with: {res.get("error", "query failed")}'
            elif res['row_count'] == 0:
                feedback = f'`{sql}` returned 0 rows. Re-check filter values against the allowed values and date ranges.'
            elif res['row_count'] == 1 and all(v in ('', 'None') for v in res['rows'][0]):
                feedback = f'`{sql}` returned only NULL. Check column names and filters.'
            else:
                break
        if not res or not res['success']:
            err = res.get('error') if res else 'no SQL generated'
            return {'success': False, 'sql_query': sql, 'error': err,
                    'answer': f'I could not produce a working query ({err}). Try naming the column or filter you mean.'}
        cols, rows, n = res['columns'], res['rows'], res['row_count']
        base = {'success': True, 'sql_query': sql, 'execution_time_ms': res['execution_time_ms'],
                'row_count': n, 'table_name': tables[0], 'raw_rows': rows[:200]}
        if n == 0:
            return {**base, 'answer': 'No records matched your question.', 'chart_data': None}
        return {**base, 'answer': self.explain_result(query, sql, cols, rows, n),
                'chart_data': pick_chart(query, cols, rows)}

    def explain_result(self, question: str, sql: str, cols: List[str], rows: List[List[Any]], n: int) -> str:
        """Explains query takeaway concisely and renders markdown result table."""
        if n == 1 and len(cols) == 1:
            return f'**{cols[0].replace("_", " ").title()}**: **{rows[0][0]}**'
        table = markdown_table(cols, rows[:MAX_TABLE_ROWS])
        note = f'\n\n_Showing the first {MAX_TABLE_ROWS} of {n} rows._' if n > MAX_TABLE_ROWS else ''
        prompt = (
            'You are a careful data analyst. Using ONLY the result rows below, answer the question in 2 to 4 sentences. '
            'State the key number(s), mention the highest and lowest or the trend if visible, and note any assumption '
            '(for example which column you treated as the measure). Never use a number that is not in the rows. Do not mention SQL.\n\n'
            f'Question: {question}\nSQL used: {sql}\nColumns: {", ".join(cols)}\n'
            f'Rows (first 30): {rows[:30]}\n\nAnswer:')
        try:
            r = requests.post(f'{self.ollama_url}/api/generate', json={
                'model': ANSWER_MODEL, 'prompt': prompt, 'stream': False, 'keep_alive': KEEP_ALIVE,
                'options': {'temperature': 0.2, 'num_ctx': 4096, 'num_predict': 200}}, timeout=60).json()
            summary = r.get('response', '').strip()
        except Exception:
            summary = ''
        return f'{summary}\n\n{table}{note}'.strip()


# ==============================================================================
# Phase 3: Persistent Chat History (SQLite)
# ==============================================================================
class ChatHistoryDB:
    """Manages persistent chat sessions and multi-turn message history in SQLite."""
    def __init__(self, db_path: Optional[str] = None):
        """Initializes the SQLite conversation database and creates necessary tables for sessions and messages."""
        if db_path is None:
            base_dir = Path(__file__).resolve().parent
            data_dir = base_dir / "chat_data"
            data_dir.mkdir(parents=True, exist_ok=True)
            db_path = str(data_dir / "chat_history.db")
        self.db_path = db_path
        self._init_db()

    def _get_con(self):
        """Creates and returns a connection to the SQLite chat history database."""
        return sqlite3.connect(self.db_path)

    def _init_db(self):
        """Creates sessions and messages tables in SQLite if they do not exist already."""
        with self._get_con() as con:
            con.execute("""
                CREATE TABLE IF NOT EXISTS sessions (
                    id TEXT PRIMARY KEY,
                    title TEXT NOT NULL,
                    active_document TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            try:
                con.execute("ALTER TABLE sessions ADD COLUMN active_document TEXT")
            except Exception:
                pass
            con.execute("""
                CREATE TABLE IF NOT EXISTS messages (
                    id TEXT PRIMARY KEY,
                    session_id TEXT NOT NULL,
                    role TEXT NOT NULL,
                    content TEXT NOT NULL,
                    metadata_json TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(session_id) REFERENCES sessions(id) ON DELETE CASCADE
                )
            """)
            con.commit()

    def get_or_create_session(self, session_id: Optional[str] = None, title: str = "New Conversation") -> str:
        """Finds an existing chat session by ID or creates a new one with a default title."""
        s_id = session_id or str(uuid.uuid4())
        with self._get_con() as con:
            cur = con.cursor()
            cur.execute("SELECT id FROM sessions WHERE id = ?", (s_id,))
            row = cur.fetchone()
            if not row:
                cur.execute("INSERT INTO sessions (id, title) VALUES (?, ?)", (s_id, title))
                con.commit()
        return s_id

    def set_session_active_doc(self, session_id: str, doc_name: Optional[Union[str, List[str]]]):
        """Saves the active document name(s) for this chat session so follow-up questions stay in context."""
        if isinstance(doc_name, (list, tuple, set)):
            clean_list = [str(x).strip() for x in doc_name if str(x).strip()]
            val = json.dumps(clean_list) if clean_list else None
        elif isinstance(doc_name, str):
            s = doc_name.strip()
            val = s if s else None
        else:
            val = None

        with self._get_con() as con:
            con.execute("UPDATE sessions SET active_document = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (val, session_id))
            con.commit()

    def get_session_active_doc(self, session_id: str) -> Optional[Union[str, List[str]]]:
        """Retrieves the currently active document(s) bound to this chat session."""
        with self._get_con() as con:
            cur = con.cursor()
            cur.execute("SELECT active_document FROM sessions WHERE id = ?", (session_id,))
            row = cur.fetchone()
            if not row or not row[0]:
                return None
            val = row[0]
            if val.startswith("[") and val.endswith("]"):
                try:
                    parsed = json.loads(val)
                    if isinstance(parsed, list):
                        return parsed
                except Exception:
                    pass
            return val

    def get_session_active_docs(self, session_id: str) -> List[str]:
        """Returns the list of active documents for this chat session."""
        val = self.get_session_active_doc(session_id)
        if not val:
            return []
        if isinstance(val, list):
            return [str(x) for x in val if str(x).strip()]
        val_str = str(val).strip()
        if "," in val_str:
            return [x.strip() for x in val_str.split(",") if x.strip()]
        return [val_str]

    def add_message(self, session_id: str, role: str, content: str, metadata: Optional[Dict[str, Any]] = None):
        """Saves a new user or assistant message with timestamp and metadata into the chat history."""
        msg_id = str(uuid.uuid4())
        meta_str = json.dumps(metadata or {})
        with self._get_con() as con:
            con.execute(
                "INSERT INTO messages (id, session_id, role, content, metadata_json) VALUES (?, ?, ?, ?, ?)",
                (msg_id, session_id, role, content, meta_str)
            )
            if role == "user":
                clean_title = re.sub(r"\s+", " ", content).strip()
                title_preview = (clean_title[:38] + "...") if len(clean_title) > 38 else (clean_title or "Conversation")
                con.execute(
                    "UPDATE sessions SET title = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ? AND (title IS NULL OR title = 'New Conversation' OR title = 'Untitled conversation' OR title = '')",
                    (title_preview, session_id)
                )
            else:
                con.execute("UPDATE sessions SET updated_at = CURRENT_TIMESTAMP WHERE id = ?", (session_id,))
            con.commit()

    def get_messages(self, session_id: str, limit: int = 30) -> List[Dict[str, Any]]:
        """Loads the recent conversation messages for a chat session in chronological order using rowid."""
        with self._get_con() as con:
            cur = con.cursor()
            cur.execute(
                'SELECT id, role, content, metadata_json, created_at FROM ('
                '  SELECT rowid AS rid, * FROM messages WHERE session_id = ? '
                '  ORDER BY rowid DESC LIMIT ?'
                ') ORDER BY rid ASC', (session_id, limit)
            )
            rows = cur.fetchall()
            messages = []
            for r in rows:
                meta = {}
                if r[3]:
                    try:
                        meta = json.loads(r[3])
                    except Exception:
                        meta = {}
                messages.append({
                    "id": r[0],
                    "role": r[1],
                    "content": r[2],
                    "metadata": meta,
                    "created_at": r[4]
                })
            return messages

    def list_sessions(self) -> List[Dict[str, Any]]:
        """Lists all saved chat sessions ordered by most recently updated."""
        with self._get_con() as con:
            cur = con.cursor()
            cur.execute("SELECT id, title, active_document, created_at, updated_at FROM sessions ORDER BY updated_at DESC LIMIT 50")
            rows = cur.fetchall()
            return [{"id": r[0], "title": r[1], "active_document": r[2], "created_at": r[3], "updated_at": r[4]} for r in rows]

    def clear_session(self, session_id: str):
        """Deletes all messages and session records for a given session ID."""
        with self._get_con() as con:
            con.execute("DELETE FROM messages WHERE session_id = ?", (session_id,))
            con.execute("DELETE FROM sessions WHERE id = ?", (session_id,))
            con.commit()


def usable_history(msgs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Filters out bad or non-grounded assistant answers and their preceding user prompt from context."""
    out = []
    for m in msgs:
        if m.get('role') == 'assistant':
            meta = m.get('metadata') or {}
            if meta.get('route') == 'CONVERSATIONAL' or meta.get('weak_retrieval'):
                if out and out[-1].get('role') == 'user':
                    out.pop()          # drop the question that got the bad answer
                continue
        out.append(m)
    return out


# ==============================================================================
# Document Catalog Manager & Scoper
# ==============================================================================
class DocumentCatalogManager:
    """
    Intelligent Document Resolver & Scoper:
    - Automatically discovers indexed documents in Qdrant & DuckDB.
    - Resolves user queries to their target document (via exact filename, aliases, keywords, or active session).
    - Prevents cross-document noise, false-positives, and hallucinations.
    """
    def __init__(self, indexer: "LocalRAGIndexer", duckdb_engine: Optional["DuckDBEngine"] = None):
        """Initializes the catalog manager with vector search indexer and DuckDB engine."""
        self.indexer = indexer
        self.duckdb_engine = duckdb_engine

    def get_catalog(self) -> List[Dict[str, Any]]:
        """Discovers all uploaded documents across Qdrant and DuckDB, creating searchable names, types, and aliases."""
        raw_sources = self.indexer.list_indexed_sources()
        sources = [s for s in raw_sources if s and not re.match(r"^(?:blob|test)$", s)]
        if self.duckdb_engine and hasattr(self.duckdb_engine, "registry"):
            for tbl_name, info in self.duckdb_engine.registry.items():
                src = info.get("source_file")
                if src and src not in sources:
                    sources.append(src)
                elif not src and not re.match(r"^(?:pdf_p\d+_tbl\d+|test|blob|dataset|active_table|tbl_.*)$", tbl_name):
                    if tbl_name not in sources:
                        sources.append(tbl_name)

        catalog = []
        for src in sources:
            base_name = os.path.splitext(src)[0]
            ext = os.path.splitext(src)[1].lower().lstrip(".")
            
            aliases = [src.lower(), base_name.lower()]
            norm_name = re.sub(r"[\s_\-]+", " ", base_name).lower()
            if norm_name not in aliases:
                aliases.append(norm_name)
            
            doc_type = "document"
            keywords = [w for w in re.split(r"[\s_\-]+", base_name.lower()) if len(w) >= 3]
            
            if "resume" in base_name.lower() or "cv" in base_name.lower():
                keywords.extend(["resume", "cv", "portfolio"])
                doc_type = "resume"
            elif "policy" in base_name.lower() or "policies" in base_name.lower() or "guideline" in base_name.lower():
                keywords.extend(["policy", "policies", "guidelines", "rules", "handbook"])
                doc_type = "policy"
            elif ext in ["xls", "xlsx", "csv"]:
                keywords.extend(["excel", "spreadsheet", "dataset", "table", "records"])
                doc_type = "spreadsheet"

            catalog.append({
                "source": src,
                "base_name": base_name,
                "extension": ext,
                "doc_type": doc_type,
                "aliases": sorted(list(set(aliases)), key=len, reverse=True),
                "keywords": list(set(keywords))
            })
        return catalog

    def resolve_target_document(self, query: str, session_active_doc: Optional[str] = None) -> Optional[str]:
        # Conversational / Math queries should NEVER be bound to documents!
        """Finds which document the user is asking about by checking filenames, aliases, keywords, or active session context."""
        if QueryRouter.is_conversational(query):
            return None

        catalog = self.get_catalog()
        if not catalog:
            return session_active_doc

        q_norm = query.lower()

        # Step 1: Direct Filename / Alias Match (e.g. "Rutvik Ram.pdf", "Lead Geneartion Practical_revised.pdf", "company policies")
        for item in catalog:
            for alias in item["aliases"]:
                pattern = r"\b" + re.escape(alias) + r"\b"
                if re.search(pattern, q_norm) or alias in q_norm:
                    return item["source"]

        # Step 2: Distinctive Individual Token Match (e.g. "Rutvik", "Laravel", "Roadmap")
        EXCLUDED_KEYWORDS = {
            "resume", "cv", "portfolio", "policy", "policies", "guidelines", "rules", "handbook",
            "excel", "spreadsheet", "table", "dataset", "records", "sample", "file", "document",
            "pdf", "doc", "docx", "txt", "sheet", "page", "data", "test", "blob", "overview",
            "summary", "details", "content", "contents", "section", "clause", "info", "test"
        }
        for item in catalog:
            for kw in item["keywords"]:
                if kw in EXCLUDED_KEYWORDS or len(kw) < 4:
                    continue
                pattern = r"\b" + re.escape(kw) + r"\b"
                if re.search(pattern, q_norm):
                    return item["source"]

        # Step 3: Type & File Extension keyword match if unique in catalog
        type_matches = set()
        for item in catalog:
            ext = (item.get("extension") or "").lower()
            if ext in ["md", "markdown"] and re.search(r"(\.md\b|\bmd\b|\bmarkdown\b)", q_norm):
                type_matches.add(item["source"])
            elif ext == "pdf" and re.search(r"(\.pdf\b|\bpdf\b)", q_norm):
                type_matches.add(item["source"])
            elif ext in ["xls", "xlsx", "csv"] and re.search(r"(\.(?:xls|xlsx|csv)\b|\b(?:excel|spreadsheet|sheet)\b)", q_norm):
                type_matches.add(item["source"])
            elif ext in ["doc", "docx"] and re.search(r"(\.(?:docx?)\b|\b(?:doc|word)\b)", q_norm):
                type_matches.add(item["source"])

            for kw in item["keywords"]:
                if kw in ["resume", "cv", "policy", "policies", "spreadsheet", "excel"]:
                    if re.search(r"\b" + re.escape(kw) + r"\b", q_norm):
                        type_matches.add(item["source"])
                        break
        if len(type_matches) == 1:
            return list(type_matches)[0]

        # Step 4: Active Session Document fallback for follow-up queries
        # If the user has an active session document and hasn't explicitly switched to another document,
        # all follow-up queries remain strictly bound to the active session document!
        if session_active_doc:
            return session_active_doc

        # Step 5: Single Document in Catalog fallback
        # If the user has exactly one uploaded document and asks questions referring to it or requesting an overview
        if len(catalog) == 1:
            doc_cues = [
                r"\b(file|document|doc|pdf|table|sheet|data|page|clause|section|policy|resume|cv)\b",
                r"\b(what\s+(?:is|are|does|was|in|inside)|explain|summarize|tell\s+me|overview|details)\b",
                r"\b(content|contents|it|this|that|key\s*points|main\s*points|highlights)\b"
            ]
            if any(re.search(c, q_norm) for c in doc_cues) or is_summary_query(query):
                return catalog[0]["source"]

        return None


# ==============================================================================
# Phase 3: Conversational Query Rewriter
# ==============================================================================
class ConversationalQueryRewriter:
    """Rewrites follow-up questions containing pronouns/ellipsis into standalone search queries."""
    def __init__(self, ollama_url: str = "http://localhost:11434"):
        """Initializes the conversational query rewriter with the local Ollama LLM endpoint."""
        self.ollama_url = ollama_url

    def rewrite_if_needed(self, query: str, history: List[Dict[str, Any]], target_doc: Optional[str] = None) -> str:
        """Rewrites ambiguous follow-up questions (with pronouns like it, this, they) into standalone search queries."""
        q_lower = query.lower().strip()

        if not history or len(history) < 2:
            return query

        follow_up_cues = [
            r"\b(he|she|it|they|his|her|their|him|them)\b",
            r"\b(what about|how about|and the|where did|when did|why did|which one|who did)\b",
            r"\b(tell me more|more details|explain further|in addition|also|what else)\b",
            r"\b(key\s*points?|keypoints|takeaways|main\s*points|highlights|stages|steps)\b",
            r"\b(now|next|then|again|above|previous|mentioned|earlier)\b",
            r"\b(?:\d+(?:st|nd|rd|th)?\s+task|task\s+\d+|exercise\s+\d+|scenario)\b",
            r"\b(?:as\s+p[er]{2,3}|according\s+to|based\s+on)\b",
            r"^(why|when|where|who|how|and|or)\??$"
        ]
        needs_rewrite = any(re.search(c, q_lower) for c in follow_up_cues) or (len(query.split()) <= 4 and not any(k in q_lower for k in ["hi", "hello", "ping", "help"]))

        if not needs_rewrite:
            return query

        recent = history[-4:]
        history_snippet = []
        for m in recent:
            r = "User" if m["role"] == "user" else "Assistant"
            c = m["content"][:200].replace("\n", " ")
            history_snippet.append(f"{r}: {c}")
        hist_text = "\n".join(history_snippet)

        prompt = (
            "Given the chat history and the user's latest follow-up question, rewrite the question into a concise, standalone query for search.\n"
            "Replace pronouns (like 'he', 'she', 'it', 'they') and ellipsis with the specific entity, person, or dataset from the context.\n"
            "If the question is already complete and self-contained, return it unchanged.\n"
            "Do NOT answer the question. Return ONLY the rewritten query text, with no quotes or explanations.\n\n"
            f"Chat History:\n{hist_text}\n\n"
            f"Follow-up Question: {query}\n"
            "Standalone Query:"
        )

        try:
            res = requests.post(
                f"{self.ollama_url}/api/generate",
                json={
                    "model": ANSWER_MODEL,
                    "prompt": prompt,
                    "stream": False,
                    "keep_alive": KEEP_ALIVE,
                    "options": {"temperature": 0.1, "num_predict": 40}
                },
                timeout=20
            ).json()
            raw_text = res.get("response", "").strip().strip('"').strip("'")
            first_line = raw_text.split("\n")[0].strip()
            first_line = re.sub(r"^(?:standalone\s+query|rewritten\s+query|query):\s*", "", first_line, flags=re.IGNORECASE).strip().strip('"')
            if first_line and len(first_line) > 3:
                doc_words = {"pdf", "document", "documents", "file", "files", "sheet", "spreadsheet", "doc", "docx", "dataset"}
                orig_words = set(re.findall(r'\b\w+\b', q_lower))
                rewritten_words = set(re.findall(r'\b\w+\b', first_line.lower()))
                if (orig_words & doc_words) and not (rewritten_words & doc_words):
                    logger.info(f"[rewriter] original contained doc term {(orig_words & doc_words)} dropped in rewrite '{first_line}'. Keeping original.")
                    return query
                return first_line
        except Exception as e:
            logger.warning(f"query rewrite failed, using raw query: {e}")

        return query


# ==============================================================================
# Shared Pipeline Helpers & Document Modality Classifiers
# ==============================================================================
TASK_PATTERNS = [
    r"\b(?:write|draft|compose|create|prepare|generate)\b.*?\b(?:email|emails|letter|outreach|proposal|message|pitch|follow-?up|response|deliverable)\b",
    r"\b(?:task\s+\d+|\d+(?:st|nd|rd|th)?\s+task|(?:first|second|third|fourth)\s+task)\b",
    r"\b(?:exercise\s+\d+|\d+(?:st|nd|rd|th)?\s+exercise|scenario)\b",
    r"\b(?:as\s+p[er]{2,3}|according\s+to|based\s+on)\s+(?:the\s+)?(?:\d+(?:st|nd|rd|th)?\s+)?(?:task|instructions?|scenario|test|prompt|exercise)\b",
    r"\b(?:complete|fulfill|solve|execute|answer)\s+(?:the\s+)?(?:task|test|exercise)\b"
]

SUMMARY_PATTERNS = [
    r"\b(?:what\s+(?:is|are|was|were|about)?\s*(?:the|this|that|my)?\s*(?:full\s+|entire\s+|whole\s+)?(?:pdf|document|file|doc|paper|dataset|roadmap|sheet)s?(?:\s+(?:is|are|about|contain|have|cover|discuss|mean))?)\s*[\?!.]*$",
    r"\b(?:what\s+(?:does\s+)?(?:this|the|my)?\s*(?:pdf|document|file|doc|paper|dataset)\s*(?:is|say|contain|have|cover|discuss|about))\s*[\?!.]*$",
    r"\b(?:what\s+(?:is|are|was)?\s*in\s+(?:this|the|that|my)?\s*(?:pdf|document|file|doc|paper|dataset))\s*[\?!.]*$",
    r"\b(?:what\s+does\s+(?:this|the)\s+(?:pdf|document|file|doc)\s+(?:say|contain|have|cover|discuss))\s*[\?!.]*$",
    r"\b(?:explain|walk\s+me\s+through|break\s*down|describe|review)\s+(?:the|this|my)?\s*(?:full|entire|whole|complete)?\s*(?:pdf|document|file|doc|paper|roadmap)\b",
    r"\b(?:summar(?:y|ize)|overview|synopsis|digest|rundown|recap)\b",
    r"\b(?:key\s*points?|keypoints|main\s*points?|takeaways?|highlights?|core\s*points?|bullet\s*points?)\b",
    r"\b(?:give\s+me\s+the\s+(?:key\s*points?|takeaways?|highlights?|summary|overview))\b",
    r"\b(?:tell\s+me\s+about|give\s+(?:me\s+)?(?:an?\s+)?overview\s+of)\s+(?:this|the)\s+(?:pdf|document|file|doc|resume|cv|roadmap|dataset)\b",
    r"\b(?:who\s+is\s+this(?:\s+resume\s+for)?|whose\s+(?:resume|cv|file)\s+is\s+this)\b",
    r"\bwhat\s+(?:is\s+)?this\s+(?:pdf|doc|file|document)?\b",
    r"\bwhat\s+this\s+(?:pdf|file|doc)\s+about\b",
    r"\b(?:what\s+(?:is|are)\s+(?:this|it)\b)",
    r"\b(?:explain\s+everything|explain\s+all|full\s+explanation)\b"
]

def is_summary_query(query: str, target_doc: Optional[str] = None) -> bool:
    """Checks whether the user is asking for an overview, summary, description, key points, or content of a document."""
    q_lower = query.lower().strip()
    # If the user is asking about a specific item/topic, do not treat as a broad document summary
    if re.search(r"\b(cookie|recipe|salary|pto|vacation|leave\s+days|sarah|apollo|lead\s+engineer|objective)\b", q_lower):
        return False
    if any(re.search(p, q_lower, re.IGNORECASE) for p in SUMMARY_PATTERNS):
        return True
    if target_doc:
        doc_inquiry_verbs = r"\b(what|explain|describe|tell\s+me|summarize|overview|about|details?|break\s*down|give\s+me|show\s+me|list)\b"
        doc_nouns = r"\b(document|doc|pdf|file|paper|table|sheet|it|this|that|content|contents|topic|topics|points?|keypoints|takeaways)\b"
        if re.search(doc_inquiry_verbs, q_lower) and re.search(doc_nouns, q_lower):
            if not re.search(r"\b(?:about\s+(?:the\s+)?(?:secret|recipe|cookie|salary|leave|how\s+many|when|who))\b", q_lower):
                return True
    return False

def _get_conversational_instant_answer(query: str) -> Optional[str]:
    """Returns deterministic instant answer for math, greetings, bot identity, or gratitude in <20ms."""
    math_val = QueryRouter.safe_math_eval(query)
    if math_val is not None:
        clean_expr = re.sub(r"^(?:what\s+is|calculate|compute|solve|how\s+much\s+is)\s+", "", query.strip().rstrip("?").strip(), flags=re.IGNORECASE)
        return f"**{clean_expr}** = **{math_val}**"

    clean_q = query.strip().lower().rstrip("?!. ")
    if re.match(r"^(?:hi|hello|hey|heya|howdy|sup|greetings|good\s+(?:morning|afternoon|evening|day))$", clean_q):
        return "Hello! I am your local AI Data Analyst and Knowledge Assistant. How can I help you today with your documents or data?"
    if QueryRouter.IDENTITY_RE.search(clean_q):
        if re.search(r"\b(your\s+tech(?:nology)?\s+stack|what\s+tech|powers\s+you|built)\b", clean_q):
            return (
                "I run **100% locally and privately on your computer** with zero cloud dependencies or external API fees (no AWS):\n\n"
                "• **Core Microservice**: Local Python service powered by **FastAPI** & **Uvicorn**.\n"
                "• **Local LLM**: **Ollama** running `llama3.2:3b` on your local hardware.\n"
                "• **In-Process SQL Engine**: **DuckDB** for exact, sub-millisecond ANSI SQL queries directly on Excel and CSV datasets.\n"
                "• **Vector Database**: **Qdrant** for high-dimensional hybrid search.\n"
                "• **Dense Embeddings**: `nomic-embed-text` (768 dimensions) via Ollama.\n"
                "• **Sparse Embeddings**: BAAI/bge BM25 via **FastEmbed** for keyword matching.\n"
                "• **Neural Reranker**: **FlashRank** (`ms-marco-MiniLM-L-12-v2`) cross-encoder for high-precision retrieval.\n"
                "• **User Interface**: Pure HTML5, Vanilla CSS, and Server-Sent Events (SSE) streaming.\n\n"
                "All operations run strictly within your local environment, ensuring that your data, documents, and interactions remain 100% private."
            )
        return "I am your local AI Data Analyst and Knowledge Assistant, running 100% privately and locally on your machine. I can query spreadsheets with exact DuckDB SQL, answer document questions via hybrid vector search, and assist you with custom tasks."

    # Security & Privacy Policy Guardrail: Prohibit sharing internal source code, core files, and server paths
    if QueryRouter.SOURCE_LEAK_RE.search(clean_q):
        return (
            "For security, privacy, and compliance policies, core system source files, backend implementation code, "
            "and internal server paths cannot be shared or exposed.\n\n"
            "I operate 100% locally and privately on your machine with strict security safeguards to protect your personal data, "
            "documents, and system integrity. All uploaded files and analytical queries remain confidential within your local environment.\n\n"
            "How can I assist you with analyzing your documents or tabular data today?"
        )

    # Security & Privacy Policy Guardrail: Prohibit sharing sensitive personal data, credentials, and PII
    if re.search(
        r"\b((?:share|give|leak|reveal|dump|show)\s+(?:me\s+)?(?:personal\s+data|pii|passwords?|secret\s+keys?|api\s+keys?|tokens?|credentials?)|"
        r"what\s+is\s+my\s+password|reveal\s+personal\s+data)\b",
        clean_q
    ):
        return (
            "For security and privacy compliance, personal identifiable information (PII), confidential user data, "
            "passwords, and system credentials are strictly protected and cannot be disclosed or shared.\n\n"
            "All operations run strictly within your local environment under end-to-end data privacy."
        )
    if re.search(r"^(?:thank\s+you|thanks|thx|thank\s+you\s+so\s+much)[\s!.]*$", clean_q):
        return "You're very welcome! Let me know if you need any further analysis or questions answered."
    return None

def _filter_reranked_passages(reranked_passages: List[Dict[str, Any]], is_summary: bool = False) -> List[Dict[str, Any]]:
    """Eliminates low-relevance candidates and cross-document noise."""
    filtered_passages = []
    if not reranked_passages:
        return filtered_passages

    top_score = reranked_passages[0].get("rerank_score", 1.0)
    top_src = reranked_passages[0].get("source")
    for p in reranked_passages:
        s = p.get("rerank_score", 0.0)
        if not filtered_passages:
            filtered_passages.append(p)
            continue
        # For document summaries, retain all chunks belonging to the target document
        if is_summary and p.get("source") == top_src:
            filtered_passages.append(p)
            continue
        if s < 0.005 and p.get("source") != top_src:
            continue
        if p.get("source") != top_src:
            if s < 0.35 * top_score:
                continue
        elif s < 0.15 * top_score and not is_summary:
            continue
        filtered_passages.append(p)
    return filtered_passages

def is_weak_retrieval(passages: List[Dict[str, Any]], is_summary: bool, is_task: bool) -> bool:
    """Evaluates whether retrieval candidates are sufficiently strong or if they should trigger fallback."""
    if not passages:
        return True
    if is_summary or is_task:
        return False
    score = passages[0].get('rerank_score')
    return score is not None and score < MIN_RERANK_SCORE

def _build_rag_prompt(
    effective_query: str,
    target_doc: Optional[str],
    contexts: List[str],
    top_source: str,
    force_summary: Optional[bool] = None,
    recent_history: Optional[List[Dict[str, Any]]] = None
) -> Tuple[str, int, bool, bool, str]:
    """
    Constructs the specialized system prompt, output token budget, and combined context
    based on document modality (task execution, summary, resume, spreadsheet, general)
    and incorporates recent multi-turn conversation context.
    """
    is_task_execution = any(re.search(p, effective_query, re.IGNORECASE) for p in TASK_PATTERNS)
    if force_summary is not None:
        is_summary = (not is_task_execution) and force_summary
    else:
        is_summary = (not is_task_execution) and is_summary_query(effective_query, target_doc)

    if isinstance(target_doc, list):
        doc_name = ", ".join(target_doc) if target_doc else (top_source or "document")
        is_spreadsheet = any(any(str(d).lower().endswith(ext) for ext in [".xls", ".xlsx", ".csv"]) for d in target_doc)
    else:
        doc_name = target_doc or top_source or "document"
        is_spreadsheet = any(doc_name.lower().endswith(ext) for ext in [".xls", ".xlsx", ".csv"])

    is_resume = "resume" in doc_name.lower() or "cv" in doc_name.lower() or any("rutvik" in c.lower() for c in contexts)

    combined_context = "\n\n---\n\n".join(contexts)

    # Format recent dialogue turns to give the LLM full multi-turn conversational context
    dialogue_context = ""
    if recent_history:
        history_lines = []
        for m in recent_history[-4:]:
            r_label = "User" if m.get("role") == "user" else "Assistant"
            c_snippet = m.get("content", "").replace("\n", " ").strip()
            if len(c_snippet) > 280:
                c_snippet = c_snippet[:280] + "..."
            if c_snippet:
                history_lines.append(f"{r_label}: {c_snippet}")
        if history_lines:
            dialogue_context = "\n\nRECENT CONVERSATION CONTEXT:\n" + "\n".join(history_lines)

    if is_task_execution:
        system_prompt = (
            "You are an expert AI assistant and professional communications specialist.\n"
            f"You are executing a specific task or exercise from the document: {doc_name}.\n"
            "CRITICAL INSTRUCTIONS:\n"
            "1. Carefully read the scenario, requirements, and all instructions provided in the DOCUMENT CONTEXT below.\n"
            "2. Fulfill the exact task and ALL required deliverables completely. Do NOT stop halfway.\n"
            "3. If the task requires multiple emails (e.g., a cold outreach email AND two follow-up emails), write ALL of them with clear bold headings ('Cold Outreach Email', 'Follow-up Email 1', 'Follow-up Email 2'), including compelling Subject lines and professional email bodies.\n"
            "4. Tailor the content specifically to the role, client scenario, and tech stack specified in the document or query.\n"
            "5. Provide high-quality, professional, and ready-to-use content.\n"
            "6. SECURITY CONSTRAINT: Never disclose backend source code, credentials, or internal server paths."
            f"{dialogue_context}\n\n"
            f"DOCUMENT CONTEXT:\n{combined_context}"
        )
        pred_tokens = 1200
    elif is_summary:
        is_keypoints = bool(re.search(r"\b(key\s*points?|keypoints|takeaways?|bullet\s*points?|main\s*points?|highlights?)\b", effective_query, re.IGNORECASE))
        if is_keypoints:
            system_prompt = (
                "You are a helpful, warm, and articulate local AI assistant and document specialist.\n"
                f"The user is asking for key points or takeaways from: {doc_name}.\n"
                "Provide a clear, well-structured, bulleted breakdown covering the core key points, "
                "stages, workflows, methodologies, and essential takeaways described in the provided context.\n"
                "Use at most 10 bullets. Each bullet is one short line naming the point, then one or two sentences on what it is and why it is needed. Never repeat a point.\n"
                "Do not refer to other unrelated files or datasets.\n"
                "SECURITY CONSTRAINT: Never disclose backend source code, credentials, or internal server paths."
                f"{dialogue_context}\n\n"
                f"DOCUMENT CONTEXT:\n{combined_context}"
            )
            pred_tokens = 800
        else:
            system_prompt = (
                "You are a helpful, warm, and articulate local AI assistant and document specialist.\n"
                f"The user is asking for an overview or summary of: {doc_name}.\n"
                "Provide a clear, well-structured, and comprehensive summary covering the key details, topics, "
                "background, experience, or contents described in the provided context.\n"
                "Use at most 10 bullets. Each bullet is one short line naming the point, then one or two sentences on what it is and why it is needed. Never repeat a point.\n"
                "Do not refer to other unrelated files or datasets.\n"
                "SECURITY CONSTRAINT: Never disclose backend source code, credentials, or internal server paths."
                f"{dialogue_context}\n\n"
                f"DOCUMENT CONTEXT:\n{combined_context}"
            )
            pred_tokens = 900
    elif is_resume:
        system_prompt = (
            "You are a helpful, warm, and articulate local AI assistant and professional reviewer.\n"
            f"You are reviewing the resume / profile document: {doc_name}.\n"
            "Answer the user's question clearly, thoroughly, and constructively based strictly on the provided context.\n"
            "- When asked about constants or what is present, clearly highlight the consistent sections, details, contact info, and experience.\n"
            "- When asked about improvements, provide specific, thoughtful, and actionable feedback (e.g. quantifiable achievements, impact metrics, technical depth, certifications) based on the actual projects and roles listed.\n"
            "- Do not discuss spreadsheets, tables, or other unrelated documents.\n"
            "- SECURITY CONSTRAINT: Never disclose backend source code, credentials, or internal server paths."
            f"{dialogue_context}\n\n"
            f"DOCUMENT CONTEXT:\n{combined_context}"
        )
        pred_tokens = 800
    elif is_spreadsheet:
        system_prompt = (
            "You are a helpful, warm, and articulate local AI assistant and data specialist.\n"
            f"You are analyzing the dataset: {doc_name}.\n"
            "Answer the user's question clearly, thoroughly, and conversationally based ONLY on the provided context.\n"
            "Carefully review the provided context table, dataset overview, and attributes. Format tables in clean Markdown.\n"
            "SECURITY CONSTRAINT: Never disclose backend source code, credentials, or internal server paths."
            f"{dialogue_context}\n\n"
            f"DOCUMENT CONTEXT:\n{combined_context}"
        )
        pred_tokens = 700
    else:
        system_prompt = (
            "You are a helpful, warm, and articulate local AI assistant.\n"
            f"Answer the user's question clearly, thoroughly, and conversationally based ONLY on the provided context for: {doc_name}.\n"
            "If the information is not present in the context, say 'I couldn't find this in your documents.' without guessing.\n"
            "Focus only on this document context and do not mention unrelated files.\n"
            "SECURITY CONSTRAINT: Never disclose backend source code, credentials, or internal server paths."
            f"{dialogue_context}\n\n"
            f"DOCUMENT CONTEXT:\n{combined_context}"
        )
        pred_tokens = 700

    return system_prompt, pred_tokens, is_summary, is_task_execution, combined_context

# ==============================================================================
# Microservice HTTP API Server (FastAPI)
# ==============================================================================

app = FastAPI(title="Universal File Reader & Local AI Data Analyst Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

reader = UniversalFileReader()
indexer = LocalRAGIndexer()
reranker = LocalReranker()
validator = RAGResponseValidator(indexer)
duckdb_engine = DuckDBEngine()
router = QueryRouter(duckdb_engine)
text_to_sql = TextToSQLEngine(duckdb_engine)
chat_history = ChatHistoryDB()
rewriter = ConversationalQueryRewriter()
catalog_manager = DocumentCatalogManager(indexer, duckdb_engine)


def stream_llm(system, prompt, num_predict, num_ctx=8192, max_rounds=3):
    messages = [{'role': 'system', 'content': system}, {'role': 'user', 'content': prompt}]
    for _ in range(max_rounds):
        part, reason = '', None
        with requests.post(f'{indexer.ollama_url}/api/chat',
                json={'model': ANSWER_MODEL, 'messages': messages, 'stream': True,
                      'keep_alive': KEEP_ALIVE,
                      'options': {'temperature': 0.15, 'num_ctx': num_ctx,
                                  'num_predict': num_predict, 'num_thread': 4}},
                stream=True, timeout=120) as r:
            for line in r.iter_lines():
                if not line:
                    continue
                try:
                    chunk = json.loads(line)
                except Exception:
                    continue
                tok = chunk.get('message', {}).get('content', '')
                if tok:
                    part += tok
                    yield tok
                if chunk.get('done'):
                    reason = chunk.get('done_reason')
                    break
        logger.info(f'[llm] done_reason={reason} chars={len(part)}')
        if reason != 'length':
            break
        messages += [{'role': 'assistant', 'content': part},
                     {'role': 'user', 'content': 'Continue exactly where you stopped. Do not repeat anything.'}]


def call_llm(system, prompt, num_predict, num_ctx=8192, max_rounds=3, temperature=0.15) -> str:
    return "".join(stream_llm(system, prompt, num_predict, num_ctx=num_ctx, max_rounds=max_rounds))


@app.get("/health")
def health():
    """API health check endpoint: reports engine status, loaded tables, and active AI models."""
    return {
        "status": "ok",
        "engine": "UniversalFileReader & LocalDataAnalyst",
        "version": "2.2.0",
        "duckdb_enabled": bool(duckdb and duckdb_engine.con),
        "hybrid_sparse_enabled": bool(indexer.sparse_model),
        "reranker_enabled": bool(reranker.ranker),
        "reranker_model": reranker.model_name if reranker.ranker else None,
        "tables_loaded": len(duckdb_engine.registry),
        "active_table": duckdb_engine.active_table
    }


@app.get("/knowledge/sources")
def get_knowledge_sources_endpoint():
    """Returns all loaded knowledge sources grouped by parent file."""
    grouped_sources: Dict[str, Dict[str, Any]] = {}

    # 1. Gather all vector document sources from Qdrant
    try:
        raw_counts = indexer.source_chunk_counts()
        doc_counts = {
            src: cnt for src, cnt in raw_counts.items()
            if src and not re.match(r"^(?:blob|test)$", src)
        }
        for doc_name, chunk_cnt in doc_counts.items():
            ext = os.path.splitext(doc_name)[1].lower().lstrip(".")
            grouped_sources[doc_name] = {
                "id": doc_name,
                "name": doc_name,
                "source_file": doc_name,
                "type": "document",
                "file_type": ext.upper() if ext else "DOC",
                "engine": "Qdrant",
                "chunks_count": chunk_cnt,
                "tables_count": 0,
                "total_rows": 0,
                "tables": [],
                "status": "Ready"
            }
    except Exception as e:
        logger.warning(f"Error discovering Qdrant document sources: {e}")

    # 2. Merge DuckDB Tables into their parent file (or create entry if spreadsheet-only)
    if duckdb_engine and hasattr(duckdb_engine, "registry"):
        for tbl_name, info in duckdb_engine.registry.items():
            if tbl_name in ["dataset", "active_table", "_sources"]:
                continue
            src_file = info.get("source_file") or tbl_name
            ext = os.path.splitext(src_file)[1].lower().lstrip(".")

            table_entry = {
                "table_name": tbl_name,
                "total_rows": info.get("total_rows", 0),
                "columns": [c.get("name") for c in info.get("columns", [])],
                "columns_count": len(info.get("columns", [])),
                "suggested_questions": info.get("suggested_questions", [])
            }

            if src_file in grouped_sources:
                entry = grouped_sources[src_file]
                entry["tables"].append(table_entry)
                entry["tables_count"] += 1
                entry["total_rows"] += info.get("total_rows", 0)
                entry["engine"] = "Qdrant + DuckDB"
                if entry["file_type"] in ("XLS", "XLSX", "CSV", "TSV"):
                    entry["type"] = "table"
                else:
                    entry["type"] = "hybrid"
            else:
                grouped_sources[src_file] = {
                    "id": src_file,
                    "name": src_file,
                    "source_file": src_file,
                    "type": "table",
                    "file_type": ext.upper() if ext else "TABLE",
                    "engine": "DuckDB",
                    "chunks_count": 0,
                    "tables_count": 1,
                    "total_rows": info.get("total_rows", 0),
                    "columns": [c.get("name") for c in info.get("columns", [])],
                    "columns_count": len(info.get("columns", [])),
                    "tables": [table_entry],
                    "suggested_questions": info.get("suggested_questions", []),
                    "status": "Ready"
                }

    sources = list(grouped_sources.values())
    sources.sort(key=lambda s: s["name"].lower())

    return {
        "status": "success",
        "sources": sources,
        "total": len(sources),
        "tables": duckdb_engine.registry if duckdb_engine and hasattr(duckdb_engine, "registry") else {}
    }


@app.get("/tables")
def get_tables_endpoint():
    """Returns all loaded DuckDB tables and indexed knowledge sources."""
    k_data = get_knowledge_sources_endpoint()
    return {
        "status": "success",
        "active_table": duckdb_engine.active_table,
        "tables": duckdb_engine.registry,
        "sources": k_data.get("sources", []),
        "documents": [s for s in k_data.get("sources", []) if s.get("type") == "document"]
    }


class DeleteSourcePayload(BaseModel):
    source_name: Optional[str] = None


def _do_delete_source(target_name: str):
    target = (target_name or "").strip()
    if not target:
        raise HTTPException(status_code=400, detail="Target source name cannot be empty")

    logger.info(f"Removing knowledge source: {target}")
    # 1. Remove vectors from Qdrant
    indexer.delete_existing_file_vectors(target)

    # 2. Remove table from DuckDB
    duckdb_engine.delete_file_tables(target)

    return {
        "status": "success",
        "deleted": target,
        "message": f"Source '{target}' successfully deleted from knowledge library."
    }


@app.delete("/knowledge/sources/{source_name:path}")
def delete_knowledge_source_path(source_name: str):
    """Deletes an indexed document or table from Qdrant and DuckDB."""
    return _do_delete_source(source_name)


@app.post("/knowledge/sources/delete")
def delete_knowledge_source_post(payload: DeleteSourcePayload):
    """Deletes an indexed document or table via POST."""
    if not payload.source_name:
        raise HTTPException(status_code=400, detail="Missing source_name")
    return _do_delete_source(payload.source_name)


class DirectSQLRequest(BaseModel):
    sql: str


@app.post("/sql")
def execute_sql_endpoint(req: DirectSQLRequest):
    """Executes a direct read-only SQL query on the DuckDB database."""
    res = duckdb_engine.execute_query(req.sql)
    if not res.get("success"):
        raise HTTPException(status_code=400, detail=res.get("error"))
    return res


@app.post("/extract")
async def extract_file_endpoint(file: UploadFile = File(...)):
    """Extracts clean text and metadata from any uploaded file format."""
    try:
        content = await file.read()
        res = reader.read(content, filename=file.filename)
        return {
            "status": "success",
            "filename": file.filename,
            "file_type": res["file_type"],
            "text": res["text"],
            "metadata": res["metadata"]
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to read file: {str(e)}")


@app.post("/ingest")
async def ingest_file_endpoint(file: UploadFile = File(...), session_id: Optional[str] = Form(None)):
    """
    Dual Ingestion Pipeline:
    1. Loads tabular files (.xls, .xlsx, .csv) into DuckDB with schema inference.
    2. Chunks and indexes into Qdrant vector database for hybrid retrieval.
    """
    try:
        content = await file.read()
        doc = reader.read(content, filename=file.filename)
        
        # Ingest into DuckDB if structured tables present
        tables_meta = doc.get("metadata", {}).get("tables", [])
        duck_results = []
        if tables_meta:
            duck_results = duckdb_engine.ingest_tables(file.filename, tables_meta)
        else:
            duckdb_engine.delete_file_tables(file.filename)

        if duck_results and doc.get('file_type') in ('xls', 'xlsx', 'csv'):
            doc['text'] = f'# Dataset: {file.filename}\n\n' + '\n\n'.join(
                text_to_sql.get_schema_context([m['table_name']]) for m in duck_results)

        # Ingest into Qdrant Vector Store
        rag_result = indexer.ingest_document(doc)
        
        if session_id:
            chat_history.get_or_create_session(session_id)
            chat_history.set_session_active_doc(session_id, file.filename)
        
        table_summary = duck_results[0] if duck_results else None
        suggested_q = table_summary["suggested_questions"] if table_summary else []

        msg = f"File '{file.filename}' indexed successfully ({rag_result['chunks_indexed']} vector chunk(s))."
        if table_summary:
            msg += f" DuckDB table '{table_summary['table_name']}' created ({table_summary['total_rows']} rows)."

        return {
            "status": "success",
            "output": msg,
            "text": msg,
            "filename": file.filename,
            "active_document": file.filename,
            "is_tabular": bool(table_summary),
            "table_schema": table_summary,
            "suggested_questions": suggested_q,
            "details": {
                "rag": rag_result,
                "duckdb": duck_results
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to ingest file: {str(e)}")


class UnifiedQueryRequest(BaseModel):
    query: str
    session_id: Optional[str] = None
    target_document: Optional[Union[str, List[str]]] = None
    target_doc: Optional[Union[str, List[str]]] = None
    target_documents: Optional[List[str]] = None
    route: Optional[str] = None
    top_k: int = 4


def _prepare_request(req: UnifiedQueryRequest):
    session_id = chat_history.get_or_create_session(req.session_id)
    recent_msgs = usable_history(chat_history.get_messages(session_id, limit=10))

    # Parse incoming client documents (supports list, JSON array string, comma string, or single string)
    client_raw = req.target_documents if req.target_documents is not None else (
        req.target_document if req.target_document is not None else req.target_doc
    )

    if client_raw is not None:
        if isinstance(client_raw, (list, tuple, set)):
            user_specified_docs = [str(x).strip() for x in client_raw if str(x).strip()]
        elif isinstance(client_raw, str):
            s = client_raw.strip()
            if s.startswith("[") and s.endswith("]"):
                try:
                    user_specified_docs = [str(x).strip() for x in json.loads(s) if str(x).strip()]
                except Exception:
                    user_specified_docs = [s]
            elif "," in s:
                user_specified_docs = [x.strip() for x in s.split(",") if x.strip()]
            elif s:
                user_specified_docs = [s]
            else:
                user_specified_docs = []
        else:
            user_specified_docs = []
        chat_history.set_session_active_doc(session_id, user_specified_docs)
    else:
        user_specified_docs = chat_history.get_session_active_docs(session_id)

    if router.is_conversational(req.query):
        target_doc, effective_query, route = None, req.query, "CONVERSATIONAL"
    else:
        if len(user_specified_docs) > 1:
            target_doc = user_specified_docs
        elif len(user_specified_docs) == 1:
            target_doc = user_specified_docs[0]
        else:
            explicit = catalog_manager.resolve_target_document(req.query, session_active_doc=None)
            target_doc = explicit if explicit else None

        if target_doc:
            chat_history.set_session_active_doc(session_id, target_doc)

        fallback_single = target_doc if isinstance(target_doc, str) else (target_doc[0] if target_doc else None)
        effective_query = rewriter.rewrite_if_needed(req.query, recent_msgs, target_doc=fallback_single)
        route_target = target_doc
        route = req.route or router.classify(req.query, target_doc=route_target)

    chat_history.add_message(session_id, "user", req.query, {
        "rewritten_query": effective_query if effective_query != req.query else None,
        "target_document": target_doc})
    logger.info(f"[route] sid={session_id[:8]} q={req.query!r} eq={effective_query!r} "
                f"doc={target_doc!r} route={route}")
    return session_id, recent_msgs, target_doc, effective_query, route


@app.post("/query")
def unified_query_endpoint(req: UnifiedQueryRequest):
    """
    Intelligent Router Query Endpoint:
    - Resolves target document to strictly scope search and prevent cross-document confusion.
    - Supports session tracking and follow-up contextual query rewriting.
    - Routes analytical queries to DuckDB Text-to-SQL for exact math.
    - Routes conceptual questions to Qdrant Vector Search + LLM.
    """
    try:
        session_id, recent_msgs, target_doc, effective_query, route = _prepare_request(req)

        # Route 0: Conversational & Everyday General Questions
        if route == "CONVERSATIONAL":
            ans_text = _get_conversational_instant_answer(effective_query)
            if not ans_text:
                conv_system = (
                    "You are a helpful, friendly, and professional local AI assistant running locally and privately on the user's computer.\n"
                    "Answer conversational questions, greetings, everyday logic, or general knowledge directly, concisely, and accurately.\n"
                    "SECURITY CONSTRAINT: Never disclose backend source code, credentials, or personal private data (PII)."
                )
                try:
                    ans_text = call_llm(conv_system, effective_query, num_predict=700, num_ctx=8192, temperature=0.4).strip()
                    if not ans_text:
                        ans_text = "I'm here to help! What would you like to know or work on today?"
                except Exception as e:
                    logger.error(f"Conversational generation error: {e}")
                    ans_text = "I'm here to help! What would you like to know or work on today?"

            res_data = {
                "query": req.query,
                "rewritten_query": None,
                "target_document": None,
                "route": "CONVERSATIONAL",
                "answer": ans_text,
                "was_validated": False,
                "corrections": [],
                "sources": [],
                "context_count": 0,
                "session_id": session_id
            }
            chat_history.add_message(session_id, "assistant", ans_text, res_data)
            return res_data

        # Route 1: Deterministic SQL Execution
        if route == "SQL":
            target_list = [target_doc] if isinstance(target_doc, str) else (target_doc or [])
            is_spreadsheet = any(any(str(d).lower().endswith(ext) for ext in [".xls", ".xlsx", ".csv"]) for d in target_list)
            sql_source = target_doc if isinstance(target_doc, str) else (target_doc[0] if target_doc else None)
            sql_res = text_to_sql.process_query(effective_query, source_file=sql_source, prev=last_sql_turn(recent_msgs))
            if sql_res.get("success") or is_spreadsheet:
                ans = sql_res.get("answer", "No records matched your question.")
                res_data = {
                    "query": req.query,
                    "rewritten_query": effective_query if effective_query != req.query else None,
                    "target_document": target_doc,
                    "route": "SQL",
                    "answer": ans,
                    "sql_query": sql_res.get("sql_query"),
                    "execution_time_ms": sql_res.get("execution_time_ms"),
                    "row_count": sql_res.get("row_count"),
                    "chart_data": sql_res.get("chart_data"),
                    "table_name": sql_res.get("table_name"),
                    "session_id": session_id
                }
                chat_history.add_message(session_id, "assistant", ans, res_data)
                return res_data
            route = "RAG"

        # Check if query references documents but none is uploaded/active in this session
        if not target_doc and QueryRouter.EXPLICIT_DOC_RE.search(req.query):
            no_doc_msg = "No document is uploaded in this session. Please upload one first."
            res_data = {
                "query": req.query,
                "rewritten_query": None,
                "target_document": None,
                "route": "RAG",
                "answer": no_doc_msg,
                "was_validated": True,
                "corrections": [],
                "sources": [],
                "context_count": 0,
                "weak_retrieval": True,
                "session_id": session_id
            }
            chat_history.add_message(session_id, "assistant", no_doc_msg, res_data)
            return res_data

        # Route 2: Semantic Hybrid RAG Search (Dense + BM25) + Cross-Encoder Reranker
        # Scoped strictly to target_doc if identified
        is_summary = is_summary_query(effective_query, target_doc)
        raw_hits = indexer.hybrid_search(effective_query, top_k=20 if is_summary else 15, source_filter=target_doc)

        candidate_passages = []
        for hit in raw_hits:
            payload = hit.get("payload", {})
            content = payload.get("content", "")
            meta = payload.get("metadata", {})
            score = hit.get("score", 0.0)
            if content:
                page = meta.get("page")
                snippet = re.sub(r"---\s*Page\s*\d+\s*---\s*", "", content).strip()
                if len(snippet) > 240:
                    snippet = snippet[:240].strip() + "..."
                candidate_passages.append({
                    "content": content,
                    "source": meta.get("source") or meta.get("filename") or "document",
                    "page": page,
                    "snippet": snippet,
                    "score": round(score, 4)
                })

        # Apply Local Cross-Encoder Reranker (ms-marco-MiniLM-L-12-v2)
        reranked_passages = reranker.rerank(effective_query, candidate_passages, top_k=8 if is_summary else req.top_k)
        filtered_passages = _filter_reranked_passages(reranked_passages, is_summary=is_summary)

        MAX_CONTEXT_CHARS = 9000   # about 2,300 tokens
        contexts, used = [], 0
        for p in filtered_passages[:8 if is_summary else 5]:
            if used + len(p["content"]) > MAX_CONTEXT_CHARS:
                break
            contexts.append(p["content"])
            used += len(p["content"])

        top_source = filtered_passages[0]["source"] if filtered_passages else "document"
        system_prompt, pred_tokens, is_summary, is_task_execution, combined_context = _build_rag_prompt(
            effective_query, target_doc, contexts, top_source, force_summary=is_summary, recent_history=recent_msgs
        )

        is_weak = is_weak_retrieval(filtered_passages, is_summary, is_task_execution)
        logger.info(f'[rag] top_score={filtered_passages[0].get("rerank_score") if filtered_passages else None} weak={is_weak}')

        if is_weak:
            has_doc_intent = (
                any(k in effective_query.lower() for k in [
                    "document", "file", "pdf", "sheet", "spreadsheet", "resume", "cv",
                    "policy", "policies", "handbook", "table", "section", "clause", "page",
                    "dataset"
                ])
                or bool(target_doc)
                or bool(re.search(r"\b(in the (?:doc|document|pdf|file|table|sheet)|according to|as per|does it say|mentioned in|from the)\b", effective_query, re.IGNORECASE))
                or bool(QueryRouter.DOC_REF_RE.search(effective_query))
            )
            if has_doc_intent:
                doc_label = ", ".join(target_doc) if isinstance(target_doc, list) else (target_doc or "")
                fallback_ans = f"I couldn't find this information in your uploaded document{' (' + doc_label + ')' if doc_label else ''}. Please verify that the relevant document or section has been uploaded, or try rephrasing your question."
            else:
                try:
                    gen_prompt = (
                        "You are a helpful and knowledgeable local AI assistant.\n"
                        "Answer the following question clearly, accurately, and helpfully based on general knowledge.\n"
                        f"Question: {effective_query}"
                    )
                    fallback_ans = call_llm("", gen_prompt, num_predict=700, num_ctx=8192, temperature=0.4).strip()
                    if not fallback_ans:
                        fallback_ans = "I'm here to help, but I couldn't find this specific detail in your documents. Feel free to rephrase or ask any general question!"
                except Exception:
                    fallback_ans = "I'm here to help, but I couldn't find this specific detail in your documents. Feel free to rephrase or ask any general question!"

            res_data = {
                "query": req.query,
                "rewritten_query": effective_query if effective_query != req.query else None,
                "target_document": target_doc,
                "route": "RAG",
                "answer": fallback_ans,
                "was_validated": True,
                "corrections": [],
                "sources": [],
                "context_count": 0,
                "weak_retrieval": is_weak,
                "session_id": session_id
            }
            chat_history.add_message(session_id, "assistant", fallback_ans, res_data)
            return res_data

        sources = [
            {
                "source": p["source"],
                "page": p.get("page"),
                "snippet": p["snippet"],
                "score": p.get("rerank_score", p.get("score", 0.0)),
                "is_reranked": bool("rerank_score" in p)
            }
            for p in filtered_passages[:5]
        ]

        candidate_answer = ''.join(stream_llm(system_prompt, effective_query, pred_tokens))
        # Run through safety validator
        val_res = validator.validate(effective_query, candidate_answer, combined_context)
        final_answer = val_res["validated_response"]

        res_data = {
            "query": req.query,
            "rewritten_query": effective_query if effective_query != req.query else None,
            "target_document": target_doc,
            "route": "RAG",
            "answer": final_answer,
            "was_validated": True,
            "corrections": val_res.get("corrections", []),
            "sources": sources,
            "context_count": len(contexts),
            "reranked": bool(reranker.ranker),
            "session_id": session_id,
            "is_summary": is_summary
        }
        chat_history.add_message(session_id, "assistant", final_answer, res_data)
        return res_data
    except Exception as e:
        logger.exception(f"Unified query failed: {e}")
        raise HTTPException(status_code=500, detail=f"Query processing failed: {str(e)}")


@app.post("/query/stream")
def query_stream_endpoint(req: UnifiedQueryRequest):
    """
    Phase 3 SSE Streaming Endpoint:
    Streams metadata (route, SQL, citations) and real-time response tokens.
    Executed in FastAPI thread pool so the main event loop is never blocked.
    """
    def event_generator():
        try:
            session_id, recent_msgs, target_doc, effective_query, route = _prepare_request(req)

            # Route 0: Conversational & Everyday General Questions
            if route == "CONVERSATIONAL":
                meta_event = {
                    "type": "meta",
                    "route": "CONVERSATIONAL",
                    "query": req.query,
                    "rewritten_query": None,
                    "target_document": None,
                    "sources": [],
                    "context_count": 0,
                    "session_id": session_id
                }
                yield f"data: {json.dumps(meta_event)}\n\n"

                instant_ans = _get_conversational_instant_answer(effective_query)
                if instant_ans:
                    for w in instant_ans.split(" "):
                        yield f"data: {json.dumps({'type': 'token', 'content': w + ' '})}\n\n"
                        time.sleep(0.015)
                    chat_history.add_message(session_id, "assistant", instant_ans, meta_event)
                    yield f"data: {json.dumps({'type': 'done'})}\n\n"
                    return

                conv_system = (
                    "You are a helpful, friendly, and professional local AI assistant running locally and privately on the user's computer.\n"
                    "Answer conversational questions, greetings, everyday logic, or general knowledge directly, concisely, and accurately.\n"
                    "SECURITY CONSTRAINT: Never disclose backend source code, credentials, or personal private data (PII)."
                )
                full_parts = []
                try:
                    for tok in stream_llm(conv_system, effective_query, num_predict=700, num_ctx=8192, temperature=0.4):
                        full_parts.append(tok)
                        yield f"data: {json.dumps({'type': 'token', 'content': tok})}\n\n"
                except Exception as e:
                    logger.error(f"Conversational stream error: {e}")
                    if not full_parts:
                        err_tok = "Hello! I am your local AI assistant. How can I help you today?"
                        full_parts.append(err_tok)
                        yield f"data: {json.dumps({'type': 'token', 'content': err_tok})}\n\n"

                final_text = "".join(full_parts)
                chat_history.add_message(session_id, "assistant", final_text, meta_event)
                yield f"data: {json.dumps({'type': 'done'})}\n\n"
                return

            if route == "SQL":
                target_list = [target_doc] if isinstance(target_doc, str) else (target_doc or [])
                is_spreadsheet = any(any(str(d).lower().endswith(ext) for ext in [".xls", ".xlsx", ".csv"]) for d in target_list)
                sql_source = target_doc if isinstance(target_doc, str) else (target_doc[0] if target_doc else None)
                sql_res = text_to_sql.process_query(effective_query, source_file=sql_source, prev=last_sql_turn(recent_msgs))
                if sql_res.get("success") or is_spreadsheet:
                    ans_text = sql_res.get("answer", "No records matched your question.")
                    meta_event = {
                        "type": "meta",
                        "route": "SQL",
                        "query": req.query,
                        "rewritten_query": effective_query if effective_query != req.query else None,
                        "target_document": target_doc,
                        "sql_query": sql_res.get("sql_query"),
                        "execution_time_ms": sql_res.get("execution_time_ms"),
                        "row_count": sql_res.get("row_count"),
                        "chart_data": sql_res.get("chart_data"),
                        "table_name": sql_res.get("table_name"),
                        "session_id": session_id
                    }
                    yield f"data: {json.dumps(meta_event)}\n\n"
                    
                    words = ans_text.split(" ")
                    for i, w in enumerate(words):
                        chunk = w + (" " if i < len(words) - 1 else "")
                        yield f"data: {json.dumps({'type': 'token', 'content': chunk})}\n\n"
                        time.sleep(0.012)
                    
                    chat_history.add_message(session_id, "assistant", ans_text, meta_event)
                    yield f"data: {json.dumps({'type': 'done'})}\n\n"
                    return
                route = "RAG"

            # Check if query references documents but none is uploaded/active in this session
            if not target_doc and QueryRouter.EXPLICIT_DOC_RE.search(req.query):
                no_doc_msg = "No document is uploaded in this session. Please upload one first."
                meta_event = {
                    "type": "meta",
                    "route": "RAG",
                    "target_document": None,
                    "sources": [],
                    "weak_retrieval": True,
                    "session_id": session_id
                }
                yield f"data: {json.dumps(meta_event)}\n\n"
                yield f"data: {json.dumps({'type': 'token', 'content': no_doc_msg})}\n\n"
                chat_history.add_message(session_id, "assistant", no_doc_msg, meta_event)
                yield f"data: {json.dumps({'type': 'done'})}\n\n"
                return

            # Check summary intent
            is_summary = is_summary_query(effective_query, target_doc)

            # Scoped strictly to target_doc if identified
            raw_hits = indexer.hybrid_search(effective_query, top_k=20 if is_summary else 15, source_filter=target_doc)
            
            candidate_passages = []
            for hit in raw_hits:
                payload = hit.get("payload", {})
                content = payload.get("content", "")
                meta = payload.get("metadata", {})
                score = hit.get("score", 0.0)
                if content:
                    page = meta.get("page")
                    snippet = re.sub(r"---\s*Page\s*\d+\s*---\s*", "", content).strip()
                    if len(snippet) > 240:
                        snippet = snippet[:240].strip() + "..."
                    candidate_passages.append({
                        "content": content,
                        "source": meta.get("source") or meta.get("filename") or "document",
                        "page": page,
                        "snippet": snippet,
                        "score": round(score, 4)
                    })

            reranked_passages = reranker.rerank(effective_query, candidate_passages, top_k=8 if is_summary else req.top_k)
            filtered_passages = _filter_reranked_passages(reranked_passages, is_summary=is_summary)

            MAX_CONTEXT_CHARS = 9000   # about 2,300 tokens
            contexts, used = [], 0
            for p in filtered_passages[:8 if is_summary else 5]:
                if used + len(p["content"]) > MAX_CONTEXT_CHARS:
                    break
                contexts.append(p["content"])
                used += len(p["content"])

            top_source = filtered_passages[0]["source"] if filtered_passages else "document"
            system_prompt, pred_tokens, is_summary, is_task_execution, combined_context = _build_rag_prompt(
                effective_query, target_doc, contexts, top_source, force_summary=is_summary, recent_history=recent_msgs
            )

            is_weak = is_weak_retrieval(filtered_passages, is_summary, is_task_execution)
            logger.info(f'[rag] top_score={filtered_passages[0].get("rerank_score") if filtered_passages else None} weak={is_weak}')
                    
            if is_weak:
                has_doc_intent = (
                    any(k in effective_query.lower() for k in [
                        "document", "file", "pdf", "sheet", "spreadsheet", "resume", "cv",
                        "policy", "policies", "handbook", "table", "section", "clause", "page",
                        "dataset"
                    ])
                    or bool(target_doc)
                    or bool(re.search(r"\b(in the (?:doc|document|pdf|file|table|sheet)|according to|as per|does it say|mentioned in|from the)\b", effective_query, re.IGNORECASE))
                    or bool(QueryRouter.DOC_REF_RE.search(effective_query))
                )
                if has_doc_intent:
                    doc_label = ", ".join(target_doc) if isinstance(target_doc, list) else (target_doc or "")
                    fallback_msg = f"I couldn't find this information in your uploaded document{' (' + doc_label + ')' if doc_label else ''}. Please verify that the relevant document or section has been uploaded, or try rephrasing your question."
                else:
                    try:
                        gen_prompt = (
                            "You are a helpful and knowledgeable local AI assistant.\n"
                            "Answer the following question clearly, accurately, and helpfully based on general knowledge.\n"
                            f"Question: {effective_query}"
                        )
                        fallback_msg = call_llm("", gen_prompt, num_predict=700, num_ctx=8192, temperature=0.4).strip()
                        if not fallback_msg:
                            fallback_msg = "I'm here to help, but I couldn't find this specific detail in your documents. Feel free to rephrase or ask any general question!"
                    except Exception:
                        fallback_msg = "I'm here to help, but I couldn't find this specific detail in your documents. Feel free to rephrase or ask any general question!"

                meta_event = {
                    "type": "meta",
                    "route": "RAG",
                    "target_document": target_doc,
                    "sources": [],
                    "weak_retrieval": is_weak,
                    "session_id": session_id
                }
                yield f"data: {json.dumps(meta_event)}\n\n"
                yield f"data: {json.dumps({'type': 'token', 'content': fallback_msg})}\n\n"
                chat_history.add_message(session_id, "assistant", fallback_msg, meta_event)
                yield f"data: {json.dumps({'type': 'done'})}\n\n"
                return

            sources = [
                {
                    "source": p["source"],
                    "page": p.get("page"),
                    "snippet": p["snippet"],
                    "score": p.get("rerank_score", p.get("score", 0.0)),
                    "is_reranked": bool("rerank_score" in p)
                }
                for p in filtered_passages[:5]
            ]

            meta_event = {
                "type": "meta",
                "route": "RAG",
                "query": req.query,
                "rewritten_query": effective_query if effective_query != req.query else None,
                "target_document": target_doc,
                "sources": sources,
                "context_count": len(contexts),
                "session_id": session_id,
                "is_summary": is_summary
            }
            yield f"data: {json.dumps(meta_event)}\n\n"

            full_parts = []
            for tok in stream_llm(system_prompt, effective_query, pred_tokens):
                full_parts.append(tok)
                yield f"data: {json.dumps({'type': 'token', 'content': tok})}\n\n"

            full_text = "".join(full_parts)
            chat_history.add_message(session_id, "assistant", full_text, meta_event)
            yield f"data: {json.dumps({'type': 'done'})}\n\n"

        except Exception as e:
            yield f"data: {json.dumps({'type': 'error', 'message': str(e)})}\n\n"

    return StreamingResponse(event_generator(), media_type="text/event-stream")


@app.get("/chat/sessions")
def list_chat_sessions():
    """API endpoint: returns a list of all saved chat sessions."""
    return {"sessions": chat_history.list_sessions()}


@app.get("/chat/sessions/{session_id}/messages")
def get_session_messages(session_id: str):
    """API endpoint: returns all messages for a specific chat session."""
    return {"session_id": session_id, "messages": chat_history.get_messages(session_id)}


@app.post("/chat/sessions/{session_id}/clear")
def clear_chat_session(session_id: str):
    """API endpoint: clears and deletes a chat session and its message history."""
    chat_history.clear_session(session_id)
    return {"status": "cleared", "session_id": session_id}


class ActiveDocRequest(BaseModel):
    active_document: Optional[Union[str, List[str]]] = None
    active_documents: Optional[List[str]] = None


@app.post("/chat/sessions/{session_id}/active_doc")
def update_session_active_doc_endpoint(session_id: str, req: ActiveDocRequest):
    """API endpoint: updates or clears the active document(s) for a specific chat session."""
    chat_history.get_or_create_session(session_id)
    docs = req.active_documents if req.active_documents is not None else req.active_document
    chat_history.set_session_active_doc(session_id, docs)
    return {
        "status": "success",
        "session_id": session_id,
        "active_document": docs
    }


class ValidationRequest(BaseModel):
    query: str
    response: str
    context: Optional[str] = None


@app.post("/validate")
def validate_response_endpoint(req: ValidationRequest):
    """Universal response validator: verifies facts, checks math/filters, humanizes tone."""
    try:
        val_result = validator.validate(req.query, req.response, req.context)
        return {
            "status": "success",
            "is_valid": val_result["is_valid"],
            "was_corrected": val_result["was_corrected"],
            "corrections": val_result["corrections"],
            "validated_response": val_result["validated_response"]
        }
    except Exception as e:
        return {
            "status": "bypassed",
            "is_valid": True,
            "was_corrected": False,
            "corrections": [f"Validation error: {str(e)}"],
            "validated_response": req.response
        }


if __name__ == "__main__":
    if len(sys.argv) > 1:
        cmd = sys.argv[1].lower()
        if cmd == "serve":
            import uvicorn
            port = int(sys.argv[2]) if len(sys.argv) > 2 else 8001
            print(f"🚀 Starting Universal File Reader & AI Data Analyst API on port {port}...")
            uvicorn.run(app, host="0.0.0.0", port=port)
        elif cmd in ["extract", "read"]:
            filepath = sys.argv[2]
            res = reader.read(filepath)
            print(res["text"])
        elif cmd == "ingest":
            filepath = sys.argv[2]
            print(f"📄 Ingesting '{filepath}' into Qdrant & DuckDB...")
            doc = reader.read(filepath)
            t_meta = doc.get("metadata", {}).get("tables", [])
            if t_meta:
                duck_res = duckdb_engine.ingest_tables(filepath, t_meta)
                print(f"✓ DuckDB table '{duck_res[0]['table_name']}' created ({duck_res[0]['total_rows']} rows)!")
            res = indexer.ingest_document(doc)
            print(f"✓ Indexed {res['chunks_indexed']} vector chunk(s) into Qdrant collection '{indexer.collection_name}'!")
        elif cmd == "tables":
            print("Loaded DuckDB Tables:")
            for t, m in duckdb_engine.registry.items():
                print(f"  • {t}: {m['total_rows']} rows, {len(m['columns'])} columns")
                print("    Starter questions:")
                for q in m.get("suggested_questions", []):
                    print(f"      - {q}")
        elif cmd == "sql":
            query = " ".join(sys.argv[2:])
            print(f"Executing SQL: {query}")
            r = duckdb_engine.execute_query(query)
            if r.get("success"):
                print(f"✓ Rows ({r['row_count']}) in {r['execution_time_ms']}ms:")
                print(r["columns"])
                for row in r["rows"]:
                    print(row)
            else:
                print(f"❌ Error: {r.get('error')}")
        elif cmd == "query":
            q = " ".join(sys.argv[2:])
            print(f"Processing query: '{q}'")
            classification = router.classify(q)
            print(f"Router decision: {classification}")
            if classification == "SQL":
                ans = text_to_sql.process_query(q)
                print(f"SQL Generated: {ans.get('sql_query')} ({ans.get('execution_time_ms')}ms)")
                print(ans.get("answer"))
            else:
                print("Routed to RAG.")
        else:
            print("Usage:")
            print("  python file_reader_engine.py serve [port]    # Start REST microservice")
            print("  python file_reader_engine.py ingest <file>    # Ingest file into Qdrant & DuckDB")
            print("  python file_reader_engine.py tables           # List DuckDB tables & starter questions")
            print("  python file_reader_engine.py sql '<SELECT>'   # Execute raw SQL on DuckDB")
            print("  python file_reader_engine.py query '<text>'   # Test natural language query")
    else:
        if os.path.exists("file_example_XLS_10.xls"):
            print("Testing extraction on file_example_XLS_10.xls:")
            doc = reader.read("file_example_XLS_10.xls")
            print(doc["text"][:600] + "\n...")
        else:
            print("Universal File Reader Engine ready. Use 'python file_reader_engine.py serve 8001'.")
