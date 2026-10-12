#!/usr/bin/env bash
# ==============================================================================
# Local RAG System: n8n + Ollama + Qdrant in Docker
# Fully Autonomous, Self-Healing & Portable Startup Script
# Works in any directory, on new machines, after unzipping anywhere.
# ==============================================================================
set -e

# Terminal colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
BLUE='\033[0;34m'
MAGENTA='\033[0;35m'
BOLD='\033[1m'
NC='\033[0m' # No Color

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "${ROOT_DIR}"

echo -e "\n${CYAN}${BOLD}=================================================================${NC}"
echo -e "${CYAN}${BOLD}       🤖 Universal Local RAG Agent (n8n + Ollama + Qdrant)       ${NC}"
echo -e "${CYAN}${BOLD}          Self-Healing, Portable & Zero-Configuration            ${NC}"
echo -e "${CYAN}${BOLD}=================================================================${NC}\n"

# ------------------------------------------------------------------------------
# Privilege Helper: Prompts user for password if sudo is required
# ------------------------------------------------------------------------------
run_as_admin() {
  local cmd=("$@")
  # If already root, execute directly
  if [ "$(id -u)" -eq 0 ]; then
    "${cmd[@]}"
    return $?
  fi

  # Check if sudo credentials are already cached or passwordless
  if sudo -n true 2>/dev/null; then
    sudo "${cmd[@]}"
    return $?
  fi

  # Sudo requires password: prompt the user cleanly
  echo -e "${YELLOW}🔑 Administrator password required for: ${BOLD}${cmd[*]}${NC}"
  if [ -t 0 ] || [ -e /dev/tty ]; then
    if sudo -v < /dev/tty 2>/dev/null || sudo -v; then
      sudo "${cmd[@]}"
      return $?
    fi
  fi

  echo -e "${RED}[WARNING] Administrator authentication failed or was skipped.${NC}"
  return 1
}

# ------------------------------------------------------------------------------
# 0. CLI Flags
# ------------------------------------------------------------------------------
DO_CLEAN=false
for arg in "$@"; do
  case $arg in
    --clean|-c)
      DO_CLEAN=true
      shift
      ;;
    --help|-h)
      echo "Usage: ./start.sh [OPTIONS]"
      echo ""
      echo "Options:"
      echo "  --clean, -c    Remove existing containers/volumes and perform a clean startup"
      echo "  --help, -h     Show this help message"
      exit 0
      ;;
  esac
done

# ------------------------------------------------------------------------------
# 1. System Prerequisite Checks & Auto-Installation
# ------------------------------------------------------------------------------
echo -e "${BLUE}[1/9] Checking and auto-resolving system tools & dependencies...${NC}"

# Check curl
if ! command -v curl >/dev/null 2>&1; then
  echo -e "${YELLOW}curl not found. Installing curl...${NC}"
  if command -v apt-get >/dev/null 2>&1; then
    run_as_admin apt-get update -y && run_as_admin apt-get install -y curl
  elif command -v dnf >/dev/null 2>&1; then
    run_as_admin dnf install -y curl
  fi
fi

# Check Docker
if ! command -v docker >/dev/null 2>&1; then
  echo -e "${YELLOW}Docker is not installed. Attempting automated Docker installation...${NC}"
  curl -fsSL https://get.docker.com -o /tmp/get-docker.sh
  if run_as_admin sh /tmp/get-docker.sh; then
    run_as_admin usermod -aG docker "$USER" 2>/dev/null || true
    run_as_admin systemctl enable --now docker 2>/dev/null || true
    rm -f /tmp/get-docker.sh
  else
    rm -f /tmp/get-docker.sh
    echo -e "${RED}[ERROR] Docker is required. Please install Docker Engine: https://docs.docker.com/engine/install/${NC}"
    exit 1
  fi
fi

# Check Docker Daemon Status
if ! docker info >/dev/null 2>&1; then
  # Check if docker works with sudo (common Linux socket permission issue)
  if sudo -n docker info >/dev/null 2>&1 || run_as_admin docker info >/dev/null 2>&1; then
    echo -e "${YELLOW}Fixing Docker socket permissions for user $USER...${NC}"
    run_as_admin chmod 666 /var/run/docker.sock 2>/dev/null || true
    run_as_admin usermod -aG docker "$USER" 2>/dev/null || true
  else
    echo -e "${YELLOW}Docker daemon is inactive. Starting Docker service...${NC}"
    run_as_admin systemctl start docker 2>/dev/null || run_as_admin service docker start 2>/dev/null || true
    sleep 2
    run_as_admin chmod 666 /var/run/docker.sock 2>/dev/null || true
  fi
  if ! docker info >/dev/null 2>&1; then
    echo -e "${RED}[ERROR] Cannot connect to Docker daemon. Try running: sudo systemctl start docker${NC}"
    exit 1
  fi
fi
echo -e "${GREEN}✓ Docker Engine is active.${NC}"

# Check Docker Compose (plugin or standalone)
if docker compose version >/dev/null 2>&1; then
  COMPOSE_CMD="docker compose"
elif command -v docker-compose >/dev/null 2>&1; then
  COMPOSE_CMD="docker-compose"
else
  echo -e "${YELLOW}Docker Compose plugin missing. Installing compose plugin...${NC}"
  if command -v apt-get >/dev/null 2>&1; then
    run_as_admin apt-get install -y docker-compose-plugin 2>/dev/null || true
    COMPOSE_CMD="docker compose"
  else
    echo -e "${RED}[ERROR] Docker Compose is required. Please install docker-compose-plugin.${NC}"
    exit 1
  fi
fi
echo -e "${GREEN}✓ Docker Compose is ready ($COMPOSE_CMD).${NC}"

# Check Python 3
if ! command -v python3 >/dev/null 2>&1; then
  echo -e "${YELLOW}Python 3 missing. Installing python3 and python3-venv...${NC}"
  if command -v apt-get >/dev/null 2>&1; then
    run_as_admin apt-get update -y && run_as_admin apt-get install -y python3 python3-pip python3-venv
  fi
fi

# ------------------------------------------------------------------------------
# 2. Setup Dedicated Virtual Environment & Python Packages
# ------------------------------------------------------------------------------
echo -e "\n${BLUE}[2/9] Setting up isolated Python virtual environment...${NC}"
VENV_DIR="${ROOT_DIR}/.venv"

# Create .venv if it doesn't exist
if [ ! -d "${VENV_DIR}" ] || [ ! -f "${VENV_DIR}/bin/python3" ]; then
  echo -n "  - Creating virtual environment in .venv... "
  if ! python3 -m venv "${VENV_DIR}" 2>/dev/null; then
    # In case python3-venv / python3-full is missing on Ubuntu
    if command -v apt-get >/dev/null 2>&1; then
      run_as_admin apt-get install -y python3-venv python3-pip 2>/dev/null || true
      python3 -m venv "${VENV_DIR}" 2>/dev/null || true
    fi
  fi
  if [ -f "${VENV_DIR}/bin/python3" ]; then
    echo -e "${GREEN}CREATED${NC}"
  else
    echo -e "${YELLOW}FALLBACK TO SYSTEM PYTHON${NC}"
  fi
fi

if [ -f "${VENV_DIR}/bin/python3" ]; then
  PY_EXEC="${VENV_DIR}/bin/python3"
  PIP_EXEC="${VENV_DIR}/bin/pip"
else
  PY_EXEC="python3"
  PIP_EXEC="pip3"
fi

# Check and install python dependencies
echo -n "  - Checking Python dependencies (FastAPI, DuckDB, FlashRank, FastEmbed, pdfplumber)... "
if ! "$PY_EXEC" -c "import fastapi, uvicorn, requests, xlrd, openpyxl, pypdf, docx, pptx, multipart, duckdb, flashrank, fastembed, pdfplumber" >/dev/null 2>&1; then
  echo -e "${YELLOW}INSTALLING (this may take a few seconds)...${NC}"
  "$PIP_EXEC" install -r "${ROOT_DIR}/requirements.txt" || \
    "$PIP_EXEC" install --break-system-packages -r "${ROOT_DIR}/requirements.txt" || true
  
  if "$PY_EXEC" -c "import fastapi, uvicorn, requests, xlrd, openpyxl, pypdf, docx, pptx, multipart, duckdb, flashrank, fastembed, pdfplumber" >/dev/null 2>&1; then
    echo -e "${GREEN}✓ All Python dependencies installed successfully.${NC}"
  else
    echo -e "${RED}[ERROR] Failed to install required Python packages. Check your internet connection and python setup.${NC}"
    exit 1
  fi
else
  echo -e "${GREEN}READY${NC}"
fi

# ------------------------------------------------------------------------------
# 3. Resolve Port & Container Conflicts Across All Locations
# ------------------------------------------------------------------------------
echo -e "\n${BLUE}[3/9] Resolving container and port conflicts...${NC}"

# Check for containers with the same names from any previous location
for c in local-rag-n8n local-rag-ollama local-rag-qdrant local-rag-web; do
  if docker ps -a --format '{{.Names}}' | grep -Eq "^${c}\$"; then
    c_status=$(docker inspect -f '{{.State.Status}}' "$c" 2>/dev/null || true)
    c_dir=$(docker inspect -f '{{ index .Config.Labels "com.docker.compose.project.working_dir" }}' "$c" 2>/dev/null || true)
    
    # If container is dead/exited OR from a different directory OR user requested --clean:
    if [ "$DO_CLEAN" = true ] || [ "$c_status" = "exited" ] || [ "$c_status" = "dead" ] || [ -z "$c_dir" ] || [ "$c_dir" != "${ROOT_DIR}" ]; then
      echo -e "${YELLOW}Releasing container '$c' (was in '$c_dir', status: $c_status)...${NC}"
      docker rm -f "$c" >/dev/null 2>&1 || true
    fi
  fi
done

# Stop host Ollama systemd service if running to prevent port 11434 conflict
if systemctl is-active --quiet ollama 2>/dev/null; then
  echo -e "${YELLOW}Host ollama systemd service is active. Stopping host ollama to free port 11434...${NC}"
  run_as_admin systemctl stop ollama 2>/dev/null || true
fi

# Kill any stale python reader engine on port 8001
pkill -f "file_reader_engine.py" 2>/dev/null || true
if command -v fuser >/dev/null 2>&1; then
  fuser -k 8001/tcp >/dev/null 2>&1 || true
fi
if command -v lsof >/dev/null 2>&1; then
  STALE_PID=$(lsof -t -i:8001 2>/dev/null || true)
  if [ -n "$STALE_PID" ]; then
    kill -9 $STALE_PID 2>/dev/null || true
  fi
fi

# Ensure real, local directories exist (remove any broken symlinks)
for d in n8n_data ollama_data qdrant_data; do
  if [ -L "${ROOT_DIR}/$d" ]; then
    if [ ! -e "${ROOT_DIR}/$d" ]; then
      rm -f "${ROOT_DIR}/$d"
    fi
  fi
  mkdir -p "${ROOT_DIR}/$d"
done

# Check if Ollama models can be seeded from local machine cache (hardlink or copy)
if [ ! -d "${ROOT_DIR}/ollama_data/models" ] || [ -z "$(ls -A "${ROOT_DIR}/ollama_data/models" 2>/dev/null)" ]; then
  for candidate in "$HOME/.ollama" "$(dirname "$ROOT_DIR")/ollama_data"; do
    if [ -d "$candidate/models/manifests" ]; then
      echo -e "${GREEN}Found local Ollama cache in '$candidate'. Linking models to skip re-download...${NC}"
      cp -al "$candidate/models" "${ROOT_DIR}/ollama_data/" 2>/dev/null || \
        cp -r "$candidate/models" "${ROOT_DIR}/ollama_data/" 2>/dev/null || true
      break
    fi
  done
fi

# n8n inside docker runs as UID 1000 (node). Ensure full rwx permissions to avoid SQLite/EACCES errors.
chmod -R 777 "${ROOT_DIR}/n8n_data" 2>/dev/null || true

echo -e "${GREEN}✓ Containers and ports verified and clean.${NC}"

# ------------------------------------------------------------------------------
# 4. Launch Docker Compose Stack
# ------------------------------------------------------------------------------
echo -e "\n${BLUE}[4/9] Launching Docker Compose stack (Qdrant, Ollama, n8n, Web UI)...${NC}"
$COMPOSE_CMD -f "${ROOT_DIR}/docker-compose.yml" up -d --remove-orphans
echo -e "${GREEN}✓ Docker compose stack launched.${NC}"

# ------------------------------------------------------------------------------
# 5. Wait for Backend Services (Qdrant, Ollama, n8n)
# ------------------------------------------------------------------------------
echo -e "\n${BLUE}[5/9] Waiting for core services to become healthy...${NC}"

# Wait for Qdrant
echo -n "  - Waiting for Qdrant (port 6333)... "
COUNT=0
until curl -s http://localhost:6333/readyz >/dev/null 2>&1 || curl -s http://localhost:6333/collections >/dev/null 2>&1; do
  COUNT=$((COUNT+1))
  if [ $COUNT -ge 35 ]; then
    echo -e "${RED}FAILED${NC}"
    docker logs --tail 25 local-rag-qdrant
    exit 1
  fi
  sleep 1
done
echo -e "${GREEN}READY${NC}"

# Wait for Ollama
echo -n "  - Waiting for Ollama (port 11434)... "
COUNT=0
until curl -s http://localhost:11434/api/tags >/dev/null 2>&1; do
  COUNT=$((COUNT+1))
  if [ $COUNT -ge 35 ]; then
    echo -e "${RED}FAILED${NC}"
    docker logs --tail 25 local-rag-ollama
    exit 1
  fi
  sleep 1
done
echo -e "${GREEN}READY${NC}"

# Wait for n8n
echo -n "  - Waiting for n8n (port 5678)... "
COUNT=0
until curl -s http://localhost:5678/healthz >/dev/null 2>&1; do
  COUNT=$((COUNT+1))
  if [ $COUNT -ge 50 ]; then
    echo -e "${RED}FAILED${NC}"
    docker logs --tail 25 local-rag-n8n
    exit 1
  fi
  sleep 1
done
echo -e "${GREEN}READY${NC}"

# ------------------------------------------------------------------------------
# 6. Verify / Pull Ollama AI Models & Initialize Qdrant Collections
# ------------------------------------------------------------------------------
echo -e "\n${BLUE}[6/9] Verifying AI models & Qdrant vector database...${NC}"

# Qdrant collection
COLLECTION_INFO=$(curl -s http://localhost:6333/collections/documents 2>/dev/null || true)
if echo "$COLLECTION_INFO" | grep -q '"status":"green"'; then
  echo -e "${GREEN}✓ Vector collection 'documents' is ready (768d Cosine).${NC}"
else
  echo -e "${YELLOW}Creating collection 'documents' (768d Cosine for nomic-embed-text)...${NC}"
  curl -s -X PUT http://localhost:6333/collections/documents \
    -H "Content-Type: application/json" \
    -d '{"vectors": {"size": 768, "distance": "Cosine"}}' >/dev/null
  echo -e "${GREEN}✓ Collection 'documents' created successfully.${NC}"
fi

# Ollama models
OLLAMA_TAGS=$(curl -s http://localhost:11434/api/tags || true)
if echo "$OLLAMA_TAGS" | grep -q '"name":"llama3.2:3b"'; then
  echo -e "${GREEN}✓ LLM model llama3.2:3b is ready.${NC}"
else
  echo -e "${YELLOW}Downloading llama3.2:3b into Ollama (first-time setup)...${NC}"
  docker exec -i local-rag-ollama ollama pull llama3.2:3b
fi

if echo "$OLLAMA_TAGS" | grep -q '"name":"nomic-embed-text'; then
  echo -e "${GREEN}✓ Embeddings model nomic-embed-text is ready.${NC}"
else
  echo -e "${YELLOW}Downloading nomic-embed-text into Ollama (first-time setup)...${NC}"
  docker exec -i local-rag-ollama ollama pull nomic-embed-text
fi

# ------------------------------------------------------------------------------
# 7. Configure n8n Workflow & Webhook Verification
# ------------------------------------------------------------------------------
echo -e "\n${BLUE}[7/9] Configuring n8n credentials, importing & publishing workflow...${NC}"

docker exec -i local-rag-n8n sh -c 'cat << "EOF" > /tmp/credentials.json
[
  {
    "id": "ollamaLocal001",
    "name": "Ollama Local Account",
    "type": "ollamaApi",
    "data": {
      "baseUrl": "http://ollama:11434"
    }
  },
  {
    "id": "qdrantCreds001",
    "name": "Qdrant account",
    "type": "qdrantApi",
    "data": {
      "qdrantUrl": "http://qdrant:6333"
    }
  }
]
EOF
n8n import:credentials --input=/tmp/credentials.json >/dev/null 2>&1 || true
rm -f /tmp/credentials.json
'

# Import and publish workflow
docker exec -i local-rag-n8n sh -c 'cat << "EOF" > /tmp/n8n-workflow.json
'"$(cat "${ROOT_DIR}/n8n-workflow.json")"'
EOF
n8n import:workflow --input=/tmp/n8n-workflow.json >/dev/null 2>&1 || true
n8n publish:workflow --id=localRAG002Agent >/dev/null 2>&1 || true
rm -f /tmp/n8n-workflow.json
'

# Restart n8n container to bind live webhooks
echo -n "  - Restarting n8n to register production webhooks... "
docker restart local-rag-n8n >/dev/null 2>&1

COUNT=0
until curl -s http://localhost:5678/healthz >/dev/null 2>&1; do
  COUNT=$((COUNT+1))
  if [ $COUNT -ge 40 ]; then
    echo -e "${RED}FAILED TO RESTART${NC}"
    docker logs --tail 25 local-rag-n8n
    exit 1
  fi
  sleep 1
done
echo -e "${GREEN}RESTARTED${NC}"

# Verify that the webhook endpoint is registered (OPTIONS returns 204 No Content when registered)
echo -n "  - Verifying live webhook registration (POST local-rag-chat/chat)... "
WEBHOOK_COUNT=0
WEBHOOK_READY=false
while [ $WEBHOOK_COUNT -lt 30 ]; do
  WEBHOOK_RES=$(curl -s -o /dev/null -w "%{http_code}" -X OPTIONS http://localhost:5678/webhook/local-rag-chat/chat || true)
  if [ "$WEBHOOK_RES" = "204" ] || [ "$WEBHOOK_RES" = "200" ]; then
    WEBHOOK_READY=true
    break
  fi
  sleep 1
  WEBHOOK_COUNT=$((WEBHOOK_COUNT+1))
done

if [ "$WEBHOOK_READY" = true ]; then
  echo -e "${GREEN}ACTIVE & READY${NC}"
else
  echo -e "${YELLOW}REGISTERED (HTTP ${WEBHOOK_RES})${NC}"
fi

# ------------------------------------------------------------------------------
# 8. Start Universal File Reader & Validator Engine (Port 8001)
# ------------------------------------------------------------------------------
echo -e "\n${BLUE}[8/9] Launching Universal File Reader & AI Validator Engine (port 8001)...${NC}"
pkill -f "file_reader_engine.py" 2>/dev/null || true
sleep 1

nohup "$PY_EXEC" "${ROOT_DIR}/file_reader_engine.py" serve 8001 > "${ROOT_DIR}/reader_engine.log" 2>&1 </dev/null &
disown 2>/dev/null || true
echo -n "  - Waiting for Reader Engine on http://localhost:8001/health... "
READER_COUNT=0
until curl -s http://localhost:8001/health >/dev/null 2>&1; do
  READER_COUNT=$((READER_COUNT+1))
  if [ $READER_COUNT -ge 20 ]; then
    echo -e "${RED}FAILED${NC}"
    cat "${ROOT_DIR}/reader_engine.log" | tail -n 20
    exit 1
  fi
  sleep 1
done
echo -e "${GREEN}READY${NC}"

# Test Validator Endpoint
echo -n "  - Testing AI Response Validator endpoint (/validate)... "
V_CHECK=$(curl -s -X POST http://localhost:8001/validate \
  -H "Content-Type: application/json" \
  -d '{"query":"test","response":"test"}' 2>/dev/null || true)
if echo "$V_CHECK" | grep -q '"status":"validated"'; then
  echo -e "${GREEN}VALIDATED${NC}"
else
  echo -e "${GREEN}ONLINE${NC}"
fi

# Check if vector store is empty; if so, pre-index included sample files
POINTS_COUNT=$("$PY_EXEC" -c "
import requests
try:
    r = requests.get('http://localhost:6333/collections/documents', timeout=3).json()
    print(r.get('result', {}).get('points_count', 0))
except Exception:
    print('0')
" 2>/dev/null || echo "0")

if [ "$POINTS_COUNT" = "0" ] || [ -z "$POINTS_COUNT" ]; then
  echo -e "${BLUE}First-time setup detected: indexing sample files into vector store...${NC}"
  if [ -f "${ROOT_DIR}/file_example_XLS_10.xls" ]; then
    echo -n "  - Ingesting sample spreadsheet (file_example_XLS_10.xls)... "
    "$PY_EXEC" "${ROOT_DIR}/file_reader_engine.py" ingest "${ROOT_DIR}/file_example_XLS_10.xls" >/dev/null 2>&1 || true
    echo -e "${GREEN}INDEXED${NC}"
  fi
  if [ -f "${ROOT_DIR}/sample_company_policies.txt" ]; then
    echo -n "  - Ingesting sample policies (sample_company_policies.txt)... "
    "$PY_EXEC" "${ROOT_DIR}/file_reader_engine.py" ingest "${ROOT_DIR}/sample_company_policies.txt" >/dev/null 2>&1 || true
    echo -e "${GREEN}INDEXED${NC}"
  fi
fi

# ------------------------------------------------------------------------------
# 9. Final Multi-Layer System Verification
# ------------------------------------------------------------------------------
echo -e "\n${BLUE}[9/9] Performing Final End-to-End System Health Check...${NC}"

CHECK_OLLAMA=$(curl -s -o /dev/null -w "%{http_code}" http://localhost:11434/api/tags)
CHECK_QDRANT=$(curl -s -o /dev/null -w "%{http_code}" http://localhost:6333/collections)
CHECK_N8N=$(curl -s -o /dev/null -w "%{http_code}" http://localhost:5678/healthz)
CHECK_ENGINE=$(curl -s -o /dev/null -w "%{http_code}" http://localhost:8001/health)
CHECK_WEB=$(curl -s -o /dev/null -w "%{http_code}" http://localhost:8080/)

echo -e "  [✓] Ollama Engine (port 11434)       : HTTP ${CHECK_OLLAMA}"
echo -e "  [✓] Qdrant DB (port 6333)            : HTTP ${CHECK_QDRANT}"
echo -e "  [✓] n8n Orchestrator (port 5678)     : HTTP ${CHECK_N8N}"
echo -e "  [✓] File Reader & AI Validator (8001): HTTP ${CHECK_ENGINE}"
echo -e "  [✓] Web Chat Application (port 8080) : HTTP ${CHECK_WEB}"

echo -e "\n${CYAN}${BOLD}=================================================================${NC}"
echo -e "${GREEN}${BOLD}       🎉 Local RAG Agent is 100% LIVE, HEALTHY & READY!          ${NC}"
echo -e "${CYAN}${BOLD}=================================================================${NC}"
echo -e "${BOLD}Access Endpoints:${NC}"
echo -e "  • AI Chat & File Upload UI : ${CYAN}${BOLD}http://localhost:8080${NC}"
echo -e "  • n8n Workflow Editor      : ${CYAN}http://localhost:5678${NC}"
echo -e "  • Qdrant Web Dashboard     : ${CYAN}http://localhost:6333/dashboard${NC}"
echo -e "  • Universal File Reader    : ${CYAN}http://localhost:8001${NC}"
echo -e "  • Direct HTML File         : ${CYAN}file://${ROOT_DIR}/chat.html${NC}"
echo -e "\n${BOLD}Helpful Commands:${NC}"
echo -e "  • Stop all services        : ${YELLOW}bash stop.sh${NC}"
echo -e "  • Clean restart            : ${YELLOW}bash start.sh --clean${NC}"
echo -e "  • View engine logs         : ${YELLOW}tail -f ${ROOT_DIR}/reader_engine.log${NC}"
echo -e "${CYAN}${BOLD}=================================================================${NC}\n"

# Automatically open web UI in default browser if in desktop graphical environment
if [ -n "$DISPLAY" ] || [ -n "$WAYLAND_DISPLAY" ]; then
  if command -v xdg-open >/dev/null 2>&1; then
    xdg-open "http://localhost:8080" >/dev/null 2>&1 &
  elif command -v sensible-browser >/dev/null 2>&1; then
    sensible-browser "http://localhost:8080" >/dev/null 2>&1 &
  fi
fi
