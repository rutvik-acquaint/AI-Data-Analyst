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
    else:
        test_duckdb_sql()
        test_intelligent_router()
        test_hybrid_and_reranker()
        test_phase3_features()
        test_flow_a_ingestion()
        time.sleep(2)
        check_qdrant_vectors()
        test_flow_b_qa()
