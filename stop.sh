#!/usr/bin/env bash
# ==============================================================================
# Local RAG System: Clean Shutdown Script
# Stops Docker Compose stack, Python Reader Engine, and releases all ports.
# ==============================================================================
set -e

# Terminal colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
BOLD='\033[1m'
NC='\033[0m' # No Color

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "${ROOT_DIR}"

echo -e "\n${CYAN}${BOLD}=================================================================${NC}"
echo -e "${CYAN}${BOLD}       🛑 Stopping Universal Local RAG Agent Services...         ${NC}"
echo -e "${CYAN}${BOLD}=================================================================${NC}\n"

# 1. Detect Docker Compose command
if docker compose version >/dev/null 2>&1; then
  COMPOSE_CMD="docker compose"
elif command -v docker-compose >/dev/null 2>&1; then
  COMPOSE_CMD="docker-compose"
else
  COMPOSE_CMD=""
fi

# 2. Stop Docker Compose stack
if [ -n "$COMPOSE_CMD" ] && [ -f "${ROOT_DIR}/docker-compose.yml" ]; then
  echo -n "  • Shutting down Docker Compose containers... "
  $COMPOSE_CMD -f "${ROOT_DIR}/docker-compose.yml" down --remove-orphans >/dev/null 2>&1 || true
  echo -e "${GREEN}DONE${NC}"
fi

# 3. Fallback: Stop individual containers by name if still running
for c in local-rag-web local-rag-n8n local-rag-ollama local-rag-qdrant; do
  if docker ps -q -f name="^${c}$" | grep -q .; then
    echo -n "  • Stopping container '$c'... "
    docker stop "$c" >/dev/null 2>&1 || true
    docker rm -f "$c" >/dev/null 2>&1 || true
    echo -e "${GREEN}STOPPED${NC}"
  fi
done

# 4. Stop Python File Reader Engine (port 8001)
echo -n "  • Stopping Universal File Reader & AI Validator Engine... "
pkill -f "file_reader_engine.py" 2>/dev/null || true

# Free port 8001 via fuser and lsof
if command -v fuser >/dev/null 2>&1; then
  fuser -k 8001/tcp >/dev/null 2>&1 || true
fi
if command -v lsof >/dev/null 2>&1; then
  STALE_PID=$(lsof -t -i:8001 2>/dev/null || true)
  if [ -n "$STALE_PID" ]; then
    kill -9 $STALE_PID 2>/dev/null || true
  fi
fi
echo -e "${GREEN}STOPPED${NC}"

echo -e "\n${GREEN}${BOLD}✓ All Local RAG containers, Python engines, and ports have been stopped cleanly.${NC}"
echo -e "${CYAN}To start the services again anytime, run: ${BOLD}bash start.sh${NC}\n"
