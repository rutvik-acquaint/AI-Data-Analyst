# 🤖 AI Data Analyst — Local & Private Document Assistant

> **100% Private, Offline-Capable AI Assistant Running Directly on Your Computer**  
> _No Cloud Subscriptions • No API Keys • Zero Data Leaks • 100% Free Forever_

---

## 🌟 What is This Project?

Imagine having your own **private ChatGPT and data analyst** running right on your laptop or PC:

- 📊 **Upload Excel & CSV Spreadsheets:** Ask questions in plain English (_"How many orders were placed after 2023?"_, _"What is the average salary by department?"_). The assistant writes exact database queries, performs **99% accurate mathematical calculations** (no AI guesswork), and draws interactive charts.
- 📄 **Upload PDFs, Word & PowerPoint Documents:** Ask questions, get summaries, or request tasks (_"Explain this document"_, _"Draft a professional outreach email based on Task 1"_). It quotes the exact page numbers and paragraphs it found the answers in.
- 💬 **Remembers Your Conversation:** You can ask natural follow-up questions without having to re-upload or retype the document's name every time.
- 🔒 **Total Privacy:** Everything runs locally on your machine. Your private documents, company spreadsheets, and client data **never leave your computer**.

---

## 🚀 Quick Start — Clone & Run (3 Simple Steps!)

You do not need to install complex software manually or be a programmer to run this project. A fully automated setup script handles everything for you.

### Prerequisites

Make sure **Git** is installed on your system. If you're not sure, open Terminal and type:

```bash
git --version
```

If it shows a version number (e.g., `git version 2.x.x`), you're ready. If not, install it:

```bash
# Ubuntu / Debian
sudo apt update && sudo apt install git -y

# Fedora
sudo dnf install git -y
```

### Step 1: Clone the Repository

Open Terminal and run:

```bash
git clone https://github.com/rutvik-acquaint/AI-Data-Analyst.git
```

This will download the entire project into a folder called `AI-Data-Analyst`.

### Step 2: Go Into the Project Folder

```bash
cd AI-Data-Analyst
```

### Step 3: Run the One-Click Startup Command

```bash
bash start.sh
```

_(If prompted, enter your computer login password so it can configure Docker and system permissions)._

### That's It! 🎉

The script will automatically:

1. Check and install any missing tools (Docker, Python environment, dependencies).
2. Download and prepare the local AI models (`llama3.2` and `nomic-embed-text`).
3. Start the database, AI engine, and web interface.
4. **Automatically open your browser** directly to the chat application at:

👉 **[http://localhost:8080](http://localhost:8080)**

> **💡 TIP — Copy-paste the full Quick Start in one go:**
>
> ```bash
> git clone https://github.com/rutvik-acquaint/AI-Data-Analyst.git
> cd AI-Data-Analyst
> bash start.sh
> ```

---

## 🛑 How to Stop the Project

When you are done and want to free up your computer's memory, open Terminal in this folder and run:

```bash
bash stop.sh
```

This cleanly shuts down all background services and releases all computer ports.

To start it back up anytime later, simply run `bash start.sh` again!

---

## 🔄 How to Update to the Latest Version

If someone has pushed a newer version, you can update your local copy with:

```bash
cd AI-Data-Analyst
git pull origin main
bash start.sh
```

This downloads the latest code changes and restarts the system. Your uploaded documents and chat history are preserved — they are stored in local data folders (not tracked by Git).

---

## 💡 How to Use the Chat Application

Once the web page opens at `http://localhost:8080`, here is how you can use it:

### 1. Uploading Documents & Spreadsheets

1. Click the **📎 Paperclip button** at the bottom left (or drag and drop a file into the window).
2. Select any file:
   - 📑 **PDF files** (`.pdf`)
   - 📊 **Excel spreadsheets** (`.xlsx`, `.xls`)
   - 📈 **CSV files** (`.csv`)
   - 📝 **Word documents** (`.docx`)
   - 📽️ **PowerPoint presentations** (`.pptx`)
   - 📄 **Text files** (`.txt`, `.md`, `.json`)
3. Click **Send** (or press Enter). The file will be indexed in seconds.

### 2. Exploring Spreadsheets with One Click

When you upload an Excel or CSV file:

- An interactive **Dataset Preview Card** appears, showing the columns and data types.
- **Starter Question Chips** (like `💡 How many total records?`, `💡 Average age by country`) pop up. Click any chip to run the query instantly!
- Click the **"🔍 How I got this"** accordion below the answer to see the exact SQL query and execution time (typically `< 5 milliseconds`).

### 3. Asking Follow-Up Questions (Context Retention)

You don't need to repeat the filename! Once a document is loaded, you can ask natural follow-ups:

- Turn 1: _"Explain what this document is about."_
- Turn 2: _"Now write a professional email for a Laravel developer with 7+ years of experience as per 1st task."_
- The assistant automatically understands which document you are talking about and delivers the finished work.

### 4. Everyday General Questions

You can also ask general questions without uploading any files:

- _"What is 25 × 14?"_
- _"Hello! Who are you and what can you do?"_
- _"Explain the difference between synchronous and asynchronous programming."_

---

## 🛠️ Frequently Asked Questions & Troubleshooting

### Q1: It says "Cannot connect to Docker daemon" or "Docker is inactive"

- **Solution:** Make sure Docker is running on your machine.
  - On Ubuntu/Debian: run `sudo systemctl start docker` in Terminal.
  - On desktop systems: open the **Docker Desktop** app and wait until the whale icon says "Engine running".
  - Then run `bash start.sh` again.

### Q2: How do I do a fresh or clean restart if something got stuck?

- Run the start script with the `--clean` flag:
  ```bash
  bash start.sh --clean
  ```
  This safely resets all containers and launches the system with a clean slate.

### Q3: Can I open the chat interface if port 8080 is blocked?

- Yes! You can open the interface directly without a web server by double-clicking the `chat.html` file in your file browser, or opening:
  `file:///path/to/AI-Data-Analyst/chat.html` in Chrome, Firefox, or Edge.

### Q4: Where are my uploaded files and chat history stored?

- Everything is kept inside this folder on your machine:
  - `chat_data/` — Chat history database (SQLite).
  - `qdrant_data/` — Search index for your documents.
  - `duckdb_data/` — Analytics database for your spreadsheets.
  - `ollama_data/` — Local AI model weights.
- Deleting the folder removes everything. Your data is never uploaded to the internet or shared with anyone.

### Q5: How do I test that everything is working properly?

- We have included an automated test suite. Run:
  ```bash
  .venv/bin/python3 test_rag.py
  ```
  It tests all AI components (Ollama, DuckDB SQL, Qdrant search, and Webhooks) and prints a green checkmark `✓` for each test.

### Q6: `git clone` says "Permission denied" or "Repository not found"

- Make sure you have access to the GitHub repository. Ask your team lead to add your GitHub account as a collaborator.
- If using SSH, ensure your SSH keys are set up: `ssh -T git@github.com`

---

## 🖥️ System Requirements

| Component               | Minimum                                                                | Recommended                           |
| :---------------------- | :--------------------------------------------------------------------- | :------------------------------------ |
| **Operating System**    | Linux (Ubuntu, Debian, Fedora, Mint) or Windows/Mac via WSL2           | Linux (Ubuntu 22.04+ or 24.04+)       |
| **Memory (RAM)**        | 4 GB                                                                   | 8 GB or more                          |
| **Processor (CPU)**     | 4 cores                                                                | 4–8 cores                             |
| **Disk Space**          | 4 GB free space                                                        | 10 GB free space (for models & cache) |
| **Internet Connection** | Required only during first-time setup (clone + model download ~2.3 GB) | Completely offline once set up        |

---

## 📐 Technical Architecture (For Developers & Admins)

For developers who want to inspect or integrate with the system:

```
┌────────────────────────────────────────────────────────────────────────┐
│                        USER WEB BROWSER                                │
│                     (http://localhost:8080)                            │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
                                    ▼
┌────────────────────────────────────────────────────────────────────────┐
│               FastAPI Reader Engine & Intelligent Router               │
│                        (http://localhost:8001)                         │
├───────────────────────────────────┬────────────────────────────────────┤
│  Intent Classification & Routing  │  Active Session Context Rewriter   │
└───────┬───────────────────────────┴─────────────────────────────┬──────┘
        │                                                         │
   [Exact Math]                                            [Semantic RAG]
        ▼                                                         ▼
┌────────────────────────┐                               ┌──────────────────────┐
│     DuckDB Engine      │                               │  Qdrant Hybrid Store │
│ (ANSI SQL Analytics)   │                               │ (Dense + Sparse BM25)│
│  • Sub-millisecond SQL │                               │  • Cosine Similarity │
│  • Read-only sandbox   │                               │  • FlashRank Rerank  │
└────────────────────────┘                               └──────────┬───────────┘
                                                                    │
                                                                    ▼
                                                         ┌──────────────────────┐
                                                         │     Ollama LLM       │
                                                         │    (llama3.2:3b)     │
                                                         │  Local CPU Inference │
                                                         └──────────────────────┘
```

### Port Map & Endpoints

| Service                  | Port    | Endpoint                          | Purpose                                  |
| :----------------------- | :------ | :-------------------------------- | :--------------------------------------- |
| **Web Chat Application** | `8080`  | `http://localhost:8080`           | Modern Web Chat UI                       |
| **FastAPI Core Engine**  | `8001`  | `http://localhost:8001`           | Router, File Ingest & SQL Engine         |
| **• Query / Router**     | `8001`  | `POST /query`                     | Unified classified query endpoint        |
| **• Realtime Streaming** | `8001`  | `POST /query/stream`              | Server-Sent Events (SSE) token stream    |
| **• File Ingestion**     | `8001`  | `POST /ingest`                    | Dual DuckDB table + Qdrant vector ingest |
| **• Direct SQL**         | `8001`  | `POST /sql`                       | Safe read-only DuckDB SQL executor       |
| **• Health Status**      | `8001`  | `GET /health`                     | Health & loaded tables status            |
| **n8n Orchestrator**     | `5678`  | `http://localhost:5678`           | Visual agent workflow builder            |
| **Qdrant Vector DB**     | `6333`  | `http://localhost:6333/dashboard` | Vector collection dashboard              |
| **Ollama Local AI**      | `11434` | `http://localhost:11434`          | Local LLM and embeddings inference       |

---

## 📂 Project Directory Structure

```
AI-Data-Analyst/
├── start.sh                    # 🚀 One-click startup script (auto-installs & runs everything)
├── stop.sh                     # 🛑 One-click clean shutdown script
├── chat.html                   # 💬 Modern Web Chat UI (served on port 8080)
├── file_reader_engine.py       # ⚙️ Core Python engine: Universal Reader, DuckDB, Router, RAG
├── docker-compose.yml          # 🐳 Docker stack: Qdrant, Ollama, n8n, Nginx Web
├── n8n-workflow.json           # 🔄 Production n8n agent workflow
├── requirements.txt            # 📦 Python package dependencies
├── test_rag.py                 # 🧪 Automated regression test suite
├── sample_company_policies.txt # 📄 Sample text document for testing
├── file_example_XLS_10.xls    # 📊 Sample Excel spreadsheet for testing
├── .gitignore                  # 🚫 Keeps runtime data out of version control
└── README.md                   # 📖 This file!
```

> **Note:** Folders like `qdrant_data/`, `ollama_data/`, `duckdb_data/`, `chat_data/`, and `.venv/` are **auto-created at runtime** by `start.sh`. They store your local data and models and are excluded from Git via `.gitignore`.

---

## 🧩 Tech Stack — What Powers This Project & Why

Every technology was hand-picked to keep the system **100% local, private, fast, and free**. Here is a plain-English breakdown:

| Technology                  | What It Does in This Project                                                                                                       | Why We Chose It                                                                                                                             |
| :-------------------------- | :--------------------------------------------------------------------------------------------------------------------------------- | :------------------------------------------------------------------------------------------------------------------------------------------ |
| **Python 3**                | The main programming language that glues everything together.                                                                      | Industry standard for AI/ML. Huge ecosystem of libraries, easy to read and maintain.                                                        |
| **FastAPI**                 | Runs the core engine server (port `8001`) that receives your questions and files.                                                  | Blazing fast (async), built-in auto-documentation, and handles real-time streaming responses effortlessly.                                  |
| **Ollama**                  | Runs the AI brain (`llama3.2` for answers, `nomic-embed-text` for understanding documents) directly on your CPU.                   | No cloud API keys needed. Models run 100% offline on your own hardware — your data never leaves your machine.                               |
| **DuckDB**                  | The analytics database engine. When you upload a spreadsheet, your data is loaded here and queried with real SQL.                  | In-process (no separate server), sub-millisecond query speed, and produces **exact mathematical answers** — no AI hallucination on numbers. |
| **Qdrant**                  | The vector search database. Stores "smart fingerprints" of your document paragraphs so the AI can find the most relevant passages. | Best-in-class hybrid search (dense vectors + sparse BM25 keywords), runs locally via Docker, and has a built-in visual dashboard.           |
| **FlashRank**               | A lightweight local reranker. After Qdrant finds candidate passages, FlashRank re-scores them for higher accuracy.                 | Runs entirely on CPU (no GPU needed), adds almost zero latency, and dramatically improves answer relevance.                                 |
| **n8n**                     | A visual workflow orchestration tool (port `5678`). It connects the chat UI to the engine and manages the agent pipeline.          | Drag-and-drop workflow builder, self-hosted, and makes it easy to add new automation steps without writing code.                            |
| **Docker & Docker Compose** | Packages Qdrant, Ollama, n8n, and the web server into isolated containers that start with one command.                             | Guarantees "works on every machine" — no dependency conflicts, no manual installs, and clean teardown with `stop.sh`.                       |
| **Nginx**                   | A tiny web server inside Docker that serves the chat page (`chat.html`) on port `8080`.                                            | Ultra-lightweight, battle-tested, and ensures the UI is served properly with correct headers and caching.                                   |
| **SQLite**                  | Stores your chat history and session context locally in a small file (`chat_history.db`).                                          | Zero-config, single-file database — perfect for conversation memory without needing a heavy database server.                                |
| **HTML / CSS / JavaScript** | The chat interface (`chat.html`) that you interact with in your browser.                                                           | No build tools, no frameworks, no npm installs — just open the file and it works in any modern browser.                                     |

### 🤔 Why This Combination?

Most AI assistants send your data to cloud servers (OpenAI, Google, etc.) and charge per query. This project takes a fundamentally different approach:

1. **Privacy by Architecture** — Ollama runs models locally, Qdrant stores vectors locally, DuckDB processes data locally. There is literally no outbound network call for your queries.
2. **Exact Math, Not Guesswork** — Spreadsheet questions are answered by DuckDB's SQL engine (deterministic, exact), not by the AI model (probabilistic, approximate). This is why calculations are **99%+ accurate**.
3. **Hybrid Search for Better Answers** — Qdrant combines semantic similarity (understanding meaning) with keyword matching (finding exact terms), then FlashRank reranks results. This three-stage pipeline dramatically reduces wrong or irrelevant answers.
4. **One-Command Experience** — Docker Compose bundles 4 services into a single `bash start.sh` command. You don't need to install Python packages, configure databases, or manage ports manually.
5. **Zero Recurring Cost** — No API keys, no subscriptions, no per-token billing. Once downloaded, the entire system works offline forever.

---

## 📄 License & Privacy

- **100% Open Source & Zero Cloud Dependencies.**
- Runs entirely on your hardware. No telemetry, no third-party tracking, and no subscriptions.
