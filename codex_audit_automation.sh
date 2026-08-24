#!/bin/bash
set -euo pipefail

DB="$HOME/.codex/logs_2.sqlite"
WAL="$HOME/.codex/logs_2.sqlite-wal"
SHM="$HOME/.codex/logs_2.sqlite-shm"

sample() {
  local prefix="$1"
  local ts files combo total stats top
  ts=$(date +%s)
  files=$(stat -f "%z" "$DB" "$WAL" "$SHM" 2>/dev/null | paste -sd, -)
  combo=$(awk -v a=$(stat -f %z "$DB") -v b=$(test -f "$WAL" && stat -f %z "$WAL" || echo 0) -v c=$(test -f "$SHM" && stat -f %z "$SHM" || echo 0) 'BEGIN{print a+b+c}')
  total=$(sqlite3 -readonly "$DB" "SELECT COUNT(*) FROM logs WHERE ts >= strftime('%s','now') - 86400;")
  stats=$(sqlite3 -readonly "$DB" "SELECT level||'|'||COUNT(*)||'|'||COALESCE(SUM(estimated_bytes),0) FROM logs WHERE ts >= strftime('%s','now') - 86400 GROUP BY level ORDER BY CASE level WHEN 'TRACE' THEN 1 WHEN 'DEBUG' THEN 2 WHEN 'INFO' THEN 3 WHEN 'WARN' THEN 4 WHEN 'ERROR' THEN 5 ELSE 6 END;")
  top=$(sqlite3 -readonly "$DB" "SELECT target||'|'||COUNT(*) FROM logs WHERE ts >= strftime('%s','now') - 86400 AND level='TRACE' GROUP BY target ORDER BY COUNT(*) DESC, target ASC LIMIT 10;")
  printf '%s_TIME=%s\n' "$prefix" "$ts"
  printf '%s_FILES=%s\n' "$prefix" "$files"
  printf '%s_COMBO=%s\n' "$prefix" "$combo"
  printf '%s_TOTAL=%s\n' "$prefix" "$total"
  printf '%s_STATS<<EOF\n%s\nEOF\n' "$prefix" "$stats"
  printf '%s_TOP<<EOF\n%s\nEOF\n' "$prefix" "$top"
}

sample "SNAP1"
sleep 60
sample "SNAP2"
