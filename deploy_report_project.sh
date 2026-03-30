#!/usr/bin/env bash
set -euo pipefail

SERVER_HOST="${SERVER_HOST:-YOUR_SERVER_IP}"
SERVER_USER="${SERVER_USER:-root}"
REMOTE_DIR="${REMOTE_DIR:-/opt/report_project}"
DOMAIN="${DOMAIN:-YOUR_DOMAIN}"
LOCAL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

RSYNC_SSH_OPTS="${RSYNC_SSH_OPTS:--o StrictHostKeyChecking=accept-new}"
SSH_OPTS="${SSH_OPTS:--o StrictHostKeyChecking=accept-new}"

if ! command -v rsync >/dev/null 2>&1; then
  echo "rsync 未安装"
  exit 1
fi

if ! command -v ssh >/dev/null 2>&1; then
  echo "ssh 未安装"
  exit 1
fi

echo "同步代码到 ${SERVER_USER}@${SERVER_HOST}:${REMOTE_DIR}"
rsync -av --delete \
  --exclude ".git" \
  --exclude "frontend/node_modules" \
  --exclude "frontend/dist" \
  --exclude "backend/__pycache__" \
  --exclude "*.tar" \
  -e "ssh ${RSYNC_SSH_OPTS}" \
  "${LOCAL_DIR}/" \
  "${SERVER_USER}@${SERVER_HOST}:${REMOTE_DIR}/"

echo "远端重建镜像并重启容器"
ssh ${SSH_OPTS} "${SERVER_USER}@${SERVER_HOST}" bash -lc "set -euo pipefail
cd '${REMOTE_DIR}'
if docker compose version >/dev/null 2>&1; then
  COMPOSE='docker compose'
elif command -v docker-compose >/dev/null 2>&1; then
  COMPOSE='docker-compose'
else
  echo '未检测到 docker compose（v2）或 docker-compose（v1）'
  exit 1
fi
\$COMPOSE down
\$COMPOSE up -d --build
\$COMPOSE ps
"

echo "健康检查"
ssh ${SSH_OPTS} "${SERVER_USER}@${SERVER_HOST}" bash -lc "set -euo pipefail
curl -fsS -o /dev/null -I http://127.0.0.1:8082/
curl -fsS -o /dev/null -I 'https://${DOMAIN}/'
echo 'OK'
"
