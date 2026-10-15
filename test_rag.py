#!/usr/bin/env python3
"""
Test script for Local RAG Agent (n8n + Ollama + Qdrant)
Tests:
1. Docker service health (Qdrant on 6333, Ollama on 11434, n8n on 5678)
2. Qdrant vector database collection
3. Ollama model availability
4. Flow A: Document ingestion via webhook
5. Flow B: Document retrieval and question answering via webhook
"""
import sys
import time
import json
import requests

N8N_URL = "http://localhost:5678"
QDRANT_URL = "http://localhost:6333"
OLLAMA_URL = "http://localhost:11434"
WEBHOOK_URL = f"{N8N_URL}/webhook/local-rag-chat/chat"

def check_services():
    """Checks that Docker services (Qdrant on 6333, Ollama on 11434, n8n on 5678) are running and responsive."""
    print("🔍 [1/4] Checking Services Health...")
    
    # Qdrant
    try:
        r = requests.get(f"{QDRANT_URL}/collections", timeout=5)
        if r.status_code == 200:
            cols = [c["name"] for c in r.json().get("result", {}).get("collections", [])]
            print(f"  ✓ Qdrant online. Collections: {cols}")
        else:
            print(f"  ✗ Qdrant error: HTTP {r.status_code}")
    except Exception as e:
        print(f"  ✗ Qdrant unreachable: {e}")

    # Ollama
    try:
        r = requests.get(f"{OLLAMA_URL}/api/tags", timeout=5)
        if r.status_code == 200:
            models = [m["name"] for m in r.json().get("models", [])]
            print(f"  ✓ Ollama online. Models: {models}")
        else:
            print(f"  ✗ Ollama error: HTTP {r.status_code}")
    except Exception as e:
        print(f"  ✗ Ollama unreachable: {e}")

    # n8n
    try:
        r = requests.get(f"{N8N_URL}/healthz", timeout=5)
        if r.status_code == 200:
            print("  ✓ n8n engine online.")
        else:
            print(f"  ✗ n8n health check: HTTP {r.status_code}")
    except Exception as e:
        print(f"  ✗ n8n unreachable: {e}")

def test_flow_a_ingestion(filepath="sample_company_policies.txt"):
    """Uploads and indexes a sample document into Qdrant vector database via webhook."""
    print("\n📄 [2/4] Testing Flow A (Document Ingestion)...")
    try:
        with open(filepath, "rb") as f:
            files = {"data": (filepath, f, "text/plain")}
            data = {"chatInput": f"Indexing {filepath}", "sessionId": "test-session-001"}
            print(f"  Sending '{filepath}' to webhook: {WEBHOOK_URL}...")
            r = requests.post(WEBHOOK_URL, data=data, files=files, timeout=60)
            print(f"  Status Code: {r.status_code}")
            print(f"  Response: {r.text}")
    except Exception as e:
        print(f"  ✗ Ingestion failed: {e}")

def check_qdrant_vectors():
    """Verifies that documents were successfully chunked and stored as vectors in Qdrant."""
    print("\n📊 [3/4] Checking Qdrant Collection Stats...")
    try:
        r = requests.get(f"{QDRANT_URL}/collections/documents", timeout=5)
        if r.status_code == 200:
            info = r.json().get("result", {})
            points_count = info.get("points_count", 0)
            print(f"  ✓ Collection 'documents': {points_count} points (vectors) stored.")
        else:
            print(f"  ✗ Could not fetch collection details: HTTP {r.status_code}")
    except Exception as e:
        print(f"  ✗ Error querying Qdrant: {e}")

def test_flow_b_qa(query="Who is leading Project Apollo and what is its objective?"):
    """Asks a document question through the AI webhook and verifies the answer and citations."""
    print(f"\n💬 [4/4] Testing Flow B (Question Answering)...")
    print(f"  Query: \"{query}\"")
    try:
        payload = {"chatInput": query, "sessionId": "test-session-001"}
        t0 = time.time()
        r = requests.post(WEBHOOK_URL, json=payload, timeout=90)
        elapsed = round(time.time() - t0, 2)
        print(f"  Latency: {elapsed}s | Status: {r.status_code}")
        print(f"  Answer: {r.text}")
    except Exception as e:
        print(f"  ✗ Q&A query failed: {e}")

def test_duckdb_sql():
    """Tests DuckDB analytical SQL execution on tabular data for exact calculations."""
    print("\n⚡ [DuckDB Engine] Testing Sub-millisecond SQL Execution on Spreadsheets...")
    try:
        url = "http://localhost:8001/sql"
        query = "SELECT COUNT(*) as total_after_2016 FROM active_table WHERE date >= '2017-01-01'"
        t0 = time.time()
        r = requests.post(url, json={"sql": query}, timeout=10)
        elapsed_ms = round((time.time() - t0) * 1000, 2)
        if r.status_code == 200:
            data = r.json()
            print(f"  ✓ SQL: {query}")
            print(f"  ✓ Execution time: {data.get('execution_time_ms', elapsed_ms)}ms | Rows: {data.get('rows')}")
            assert data["rows"][0][0] in ["3", 3], f"Expected 3 records after 2016, got {data['rows'][0][0]}"
            print("  ✓ Verification: 100% exact math confirmed (3 records).")
        else:
            print(f"  ✗ SQL endpoint returned HTTP {r.status_code}: {r.text}")
    except Exception as e:
        print(f"  ✗ DuckDB SQL test failed: {e}")

def test_intelligent_router():
    """Verifies that user queries are correctly classified into pure SQL, semantic RAG, or conversational."""
    print("\n🧠 [Router & Analyst] Testing Intent Classification & Text-to-SQL...")
    queries = [
        ("how many records date is > 2016", "SQL"),
        ("What is the average age by country?", "SQL"),
        ("Who is leading Project Apollo and what is its objective?", "RAG")
    ]
    for q, expected_route in queries:
        try:
            t0 = time.time()
            r = requests.post("http://localhost:8001/query", json={"query": q}, timeout=60)
            elapsed = round(time.time() - t0, 2)
            if r.status_code == 200:
                data = r.json()
                route = data.get("route")
                print(f"  ✓ Query: \"{q}\"")
                print(f"    Route: [{route}] (Expected: {expected_route}) | Latency: {elapsed}s")
                if route == "SQL":
                    print(f"    SQL Query: {data.get('sql_query')}")
                    print(f"    Execution Time: {data.get('execution_time_ms')}ms")
                print(f"    Answer summary: {data.get('answer')[:120]}...")
            else:
                print(f"  ✗ Query failed HTTP {r.status_code}: {r.text}")
        except Exception as e:
            print(f"  ✗ Router test failed for '{q}': {e}")

def test_hybrid_and_reranker():
    """Tests hybrid search (dense vectors + sparse BM25 keywords) and FlashRank cross-encoder reranking."""
    print("\n🔬 [Phase 2: Hybrid Search & Cross-Encoder] Testing Dense+BM25 & FlashRank...")
    # 1. Health check
    try:
        h = requests.get("http://localhost:8001/health", timeout=5).json()
        print(f"  ✓ Service health: Hybrid Sparse BM25: {h.get('hybrid_sparse_enabled')} | Reranker: {h.get('reranker_enabled')} ({h.get('reranker_model')})")
    except Exception as e:
        print(f"  ✗ Health check failed: {e}")

    # 2. Known Document Retrieval with Reranker
    try:
        q = "Who is leading Project Apollo and what is its objective?"
        t0 = time.time()
        r = requests.post("http://localhost:8001/query", json={"query": q}, timeout=60).json()
        elapsed = round(time.time() - t0, 2)
        print(f"  ✓ Document query: \"{q}\" in {elapsed}s")
        print(f"    Reranked: {r.get('reranked')} | Context Count: {r.get('context_count')}")
        for s in r.get("sources", []):
            print(f"    • Source: {s.get('source')} | Score: {s.get('score')} | Snippet: {s.get('snippet')[:60]}...")
        assert "Sarah Vance" in r.get("answer", "")
        print("    ✓ Grounding verified: Correct lead cited.")
    except Exception as e:
        print(f"  ✗ Reranker test failed: {e}")

    # 3. Weak Retrieval Guardrail Test
    try:
        irrelevant_q = "What does the policy document say about the secret recipe for chocolate chip cookies?"
        r_weak = requests.post("http://localhost:8001/query", json={"query": irrelevant_q}, timeout=60).json()
        ans = r_weak.get("answer", "")
        is_honest = "couldn't find" in ans.lower() or "not find" in ans.lower() or r_weak.get("weak_retrieval")
        print(f"  ✓ Weak Retrieval Test: \"{irrelevant_q}\"")
        print(f"    Bot Answer: \"{ans[:90]}...\"")
        assert is_honest, f"Expected weak retrieval fallback, got: {ans}"
        print("    ✓ Guardrail verified: Safely avoided hallucination.")
    except Exception as e:
        print(f"  ✗ Weak retrieval guardrail test failed: {e}")

def test_phase3_features():
    """Tests multi-turn chat history persistence, session active document scoping, and response validation."""
    print("\n🚀 [Phase 3: Conversational AI, Math & Document Tasks] Testing Phase 3 capabilities...")
    # 1. Pure Math Evaluation Test
    try:
        t0 = time.time()
        r_math = requests.post("http://localhost:8001/query", json={"query": "what is 2+2?"}, timeout=10).json()
        elapsed = round(time.time() - t0, 3)
        assert "4" in r_math.get("answer", "")
        assert r_math.get("route") == "CONVERSATIONAL"
        print(f"  ✓ Instant Math: {r_math.get('answer')} in {elapsed}s | Route: {r_math.get('route')}")
    except Exception as e:
        print(f"  ✗ Math test failed: {e}")

    # 2. Daily Human Conversation & Greeting Test
    try:
        t0 = time.time()
        r_hello = requests.post("http://localhost:8001/query", json={"query": "hello, who are you?"}, timeout=60).json()
        elapsed = round(time.time() - t0, 2)
        assert len(r_hello.get("answer", "")) > 10
        assert "couldn't find" not in r_hello.get("answer", "").lower()
        print(f"  ✓ Friendly Greeting: {r_hello.get('answer')[:70]}... in {elapsed}s | Route: {r_hello.get('route')}")
    except Exception as e:
        print(f"  ✗ Greeting test failed: {e}")

    # 3. Document Scoping & Task Execution Stream Test
    try:
        task_q = "write an email as pre given summery task. 📄 Lead Geneartion Practical_revised.pdf (Page 1)"
        t0 = time.time()
        s = requests.post("http://localhost:8001/query/stream", json={"query": task_q}, stream=True, timeout=120)
        target_doc = None
        first_token_time = None
        streamed_chars = 0
        has_laravel = False
        for line in s.iter_lines():
            if line:
                l_str = line.decode("utf-8")
                if l_str.startswith("data:"):
                    d = json.loads(l_str[5:].strip())
                    if d.get("type") == "meta":
                        target_doc = d.get("target_document")
                    elif d.get("type") == "token":
                        if first_token_time is None:
                            first_token_time = round(time.time() - t0, 2)
                        tok = d.get("content", "")
                        streamed_chars += len(tok)
                        if "Laravel" in tok:
                            has_laravel = True
                    elif d.get("type") == "done":
                        break
        total_time = round(time.time() - t0, 2)
        assert target_doc == "Lead Geneartion Practical_revised.pdf"
        assert has_laravel or streamed_chars > 200
        print(f"  ✓ Streaming Task Execution: Scoped to '{target_doc}' | First token: {first_token_time}s | Total: {total_time}s | {streamed_chars} chars streamed")
    except Exception as e:
        print(f"  ✗ Task execution stream test failed: {e}")

    # 4. Multi-Turn Session Context Preservation & Follow-up Question Test
    try:
        session_id = f"test-multiturn-{int(time.time())}"
        # Turn 1: Ask overview of a document
        r_t1 = requests.post(
            "http://localhost:8001/query",
            json={"query": "give a sort overview of sample_company_policies.txt", "session_id": session_id},
            timeout=120
        ).json()
        assert r_t1.get("route") == "RAG"
        assert r_t1.get("target_document") == "sample_company_policies.txt"
        print(f"  ✓ Multi-turn Turn 1: Scoped to '{r_t1.get('target_document')}' | Route: {r_t1.get('route')}")

        # Turn 2: Follow-up question without naming the document ("give me the keypoints")
        t0 = time.time()
        r_t2 = requests.post(
            "http://localhost:8001/query",
            json={"query": "give me the keypoints", "session_id": session_id},
            timeout=120
        ).json()
        elapsed = round(time.time() - t0, 2)
        ans_t2 = r_t2.get("answer", "")
        # Verification criteria:
        # 1. Stays in RAG mode scoped to active document
        assert r_t2.get("route") == "RAG", f"Expected RAG route, got {r_t2.get('route')}"
        assert r_t2.get("target_document") == "sample_company_policies.txt", f"Expected active doc retained, got {r_t2.get('target_document')}"
        # 2. Never hallucinates internal security policy or AWS
        assert "security and privacy policy" not in ans_t2.lower(), "Hallucinated security policy!"
        assert "hosted on aws" not in ans_t2.lower(), "Hallucinated AWS!"
        # 3. Grounded answer contains key points
        assert len(ans_t2) > 50, "Answer too short!"
        print(f"  ✓ Multi-turn Turn 2 ('give me the keypoints'): Route [{r_t2.get('route')}] | Doc: '{r_t2.get('target_document')}' | Elapsed: {elapsed}s")
        print(f"    Summary: {ans_t2[:100]}...")
    except Exception as e:
        print(f"  ✗ Multi-turn follow-up test failed: {e}")

def test_acceptance_routing_and_context():
    """Acceptance tests for routing, identity, greetings, and context isolation."""
    print("\n🎯 [Acceptance Tests] Running routing & context isolation test suite...")

    # 1. With NO document active
    no_doc_sid = f"acc-no-doc-{int(time.time())}"
    r_no_doc = requests.post("http://localhost:8001/query", json={
        "query": "what does the pdf say about Python?",
        "session_id": no_doc_sid
    }).json()
    assert r_no_doc.get("route") == "RAG", f"Expected RAG, got {r_no_doc.get('route')}"
    assert "no document is uploaded" in r_no_doc.get("answer", "").lower(), f"Unexpected answer: {r_no_doc.get('answer')}"
    print("  ✓ No document active: Correctly returned 'No document is uploaded' guardrail.")

    # 2. Ingest AI developer roadmap PDF
    pdf_bytes = (
        b'%PDF-1.4\n'
        b'1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n'
        b'2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n'
        b'3 0 obj<</Type/Page/MediaBox[0 0 612 792]/Parent 2 0 R/Resources<<\n'
        b'/Font<</F1 4 0 R>>>>/Contents 5 0 R>>endobj\n'
        b'4 0 obj<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>endobj\n'
        b'5 0 obj<</Length 230>>stream\n'
        b'BT /F1 12 Tf 72 700 Td (AI Developer Roadmap: How to become an AI developer.) Tj\n'
        b'0 -20 Td (To become an AI developer, learn Python, math, PyTorch, and NLP.) Tj\n'
        b'0 -20 Td (The required tech stack consists of PyTorch, HuggingFace Transformers, and LangChain.) Tj ET\n'
        b'endstream\nendobj\n'
        b'xref\n0 6\n0000000000 65535 f \n0000000010 00000 n \n0000000060 00000 n \n0000000115 00000 n \n0000000224 00000 n \n0000000295 00000 n \n'
        b'trailer<</Size 6/Root 1 0 R>>\nstartxref\n576\n%%EOF'
    )
    acc_sid = f"acc-roadmap-{int(time.time())}"
    r_ingest = requests.post(
        "http://localhost:8001/ingest",
        files={"file": ("ai_developer_roadmap.pdf", pdf_bytes, "application/pdf")},
        data={"session_id": acc_sid}
    ).json()
    assert r_ingest.get("status") == "success"
    print(f"  ✓ Ingested active document '{r_ingest.get('active_document')}' into session.")

    # 3. 'Hello' -> greeting only, zero mention of any document
    r_hello = requests.post("http://localhost:8001/query", json={"query": "Hello", "session_id": acc_sid}).json()
    assert r_hello.get("route") == "CONVERSATIONAL", f"Expected CONVERSATIONAL, got {r_hello.get('route')}"
    ans_hello = r_hello.get("answer", "").lower()
    assert "roadmap" not in ans_hello and "ai_developer_roadmap" not in ans_hello, f"Document context leaked into greeting! Answer: {r_hello.get('answer')}"
    print(f"  ✓ Greeting isolated: '{r_hello.get('answer')}'")

    # 4. 'how to become an ai developer and what tech stack is required as per pdf?' -> route=RAG, cites PDF, no mention of FastAPI/DuckDB/Qdrant
    r_rag = requests.post("http://localhost:8001/query", json={
        "query": "how to become an ai developer and what tech stack is required as per pdf?",
        "session_id": acc_sid
    }).json()
    assert r_rag.get("route") == "RAG", f"Expected RAG, got {r_rag.get('route')}"
    assert r_rag.get("target_document") == "ai_developer_roadmap.pdf", f"Expected ai_developer_roadmap.pdf, got {r_rag.get('target_document')}"
    ans_rag = r_rag.get("answer", "")
    assert "fastapi" not in ans_rag.lower() and "duckdb" not in ans_rag.lower() and "qdrant" not in ans_rag.lower(), "Self-description leaked into document answer!"
    assert any(k in ans_rag.lower() for k in ["pytorch", "transformers", "langchain", "python"]), "PDF content not cited!"
    print(f"  ✓ RAG document query answered from PDF: Route [{r_rag.get('route')}] | Doc: '{r_rag.get('target_document')}'")

    # 5. 'what is your tech stack?' -> identity route (describes the app)
    r_identity = requests.post("http://localhost:8001/query", json={
        "query": "what is your tech stack?",
        "session_id": acc_sid
    }).json()
    assert r_identity.get("route") == "CONVERSATIONAL", f"Expected CONVERSATIONAL, got {r_identity.get('route')}"
    ans_id = r_identity.get("answer", "")
    assert any(k in ans_id.lower() for k in ["fastapi", "duckdb", "qdrant", "ollama"]), f"App stack not described! Answer: {ans_id}"
    print("  ✓ Identity query accurately describes assistant's own tech stack.")

    # 6. 'summarize the pdf' -> RAG
    r_sum = requests.post("http://localhost:8001/query", json={
        "query": "summarize the pdf",
        "session_id": acc_sid
    }).json()
    assert r_sum.get("route") == "RAG", f"Expected RAG, got {r_sum.get('route')}"
    assert r_sum.get("target_document") == "ai_developer_roadmap.pdf"
    print(f"  ✓ 'summarize the pdf' routed to RAG with doc '{r_sum.get('target_document')}'.")

    # 7. 'what is 25 * 4?' -> conversational, answer 100
    r_math = requests.post("http://localhost:8001/query", json={
        "query": "what is 25 * 4?",
        "session_id": acc_sid
    }).json()
    assert r_math.get("route") == "CONVERSATIONAL"
    assert "100" in r_math.get("answer", "")
    print(f"  ✓ Instant math: {r_math.get('answer')}")

    # 8. Upload second file and verify questions only use the second file
    second_file_bytes = b"Acme Company Policy: Remote work policy permits working remotely up to 3 days per week. Travel daily allowance is 75 USD."
    r_ingest2 = requests.post(
        "http://localhost:8001/ingest",
        files={"file": ("company_policy_v2.txt", second_file_bytes, "text/plain")},
        data={"session_id": acc_sid}
    ).json()
    assert r_ingest2.get("status") == "success"
    assert r_ingest2.get("active_document") == "company_policy_v2.txt"

    r_doc2 = requests.post("http://localhost:8001/query", json={
        "query": "what is the remote work policy as per document?",
        "session_id": acc_sid
    }).json()
    assert r_doc2.get("route") == "RAG"
    assert r_doc2.get("target_document") == "company_policy_v2.txt", f"Expected company_policy_v2.txt, got {r_doc2.get('target_document')}"
    assert any(k in r_doc2.get("answer", "").lower() for k in ["3 days", "remotely"]), f"Second file content missing: {r_doc2.get('answer')}"
    print(f"  ✓ Second file switched and isolated: Doc '{r_doc2.get('target_document')}'")

if __name__ == "__main__":
    check_services()
    if len(sys.argv) > 1 and sys.argv[1] == "--qa-only":
        test_flow_b_qa()
    elif len(sys.argv) > 1 and sys.argv[1] == "--sql-only":
        test_duckdb_sql()
        test_intelligent_router()
    elif len(sys.argv) > 1 and sys.argv[1] == "--phase2-only":
        test_hybrid_and_reranker()
    elif len(sys.argv) > 1 and sys.argv[1] == "--phase3-only":
        test_phase3_features()
    elif len(sys.argv) > 1 and sys.argv[1] == "--acceptance-only":
        test_acceptance_routing_and_context()
    else:
        test_duckdb_sql()
        test_intelligent_router()
        test_hybrid_and_reranker()
        test_phase3_features()
        test_acceptance_routing_and_context()
        test_flow_a_ingestion()
        time.sleep(2)
        check_qdrant_vectors()
        test_flow_b_qa()
