#!/bin/bash
# Phase 3.1 — 每小时同步坐席日报(悦达 3 班 + 远程日报)
# crontab: 0 * * * * /opt/bi-backend/cron_sync_agents.sh
echo "=== $(date '+%Y-%m-%d %H:%M:%S') sync start ==="
curl -sS -X POST http://127.0.0.1:5000/api/agents/sync \
     -H "X-User-Role: superadmin" \
     --max-time 60
echo
echo "=== $(date '+%Y-%m-%d %H:%M:%S') sync end ==="
