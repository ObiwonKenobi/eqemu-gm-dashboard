#!/bin/bash
# ==============================================================
# EQEmu GM Command Report — Export Script
# Project: https://github.com/YOUR_USERNAME/eqemu-gm-dashboard
# Version: 1.2  |  Created: May 2026
#
# Runs as a persistent daemon, regenerating the GM report every
# 60 seconds. Set Unraid User Scripts to "At Startup of Array"
# (no cron schedule needed).
#
# v1.2 — Performance: parallel MySQL queries (~4-5s saved) and
# a persistent worker container (~8-15s saved per cycle) instead
# of docker run --rm on every tick.
#
# Free to use and adapt. Please keep this attribution intact.
# ==============================================================

# ── Configuration — override any of these in your environment or .env ────────
# Source your .env file if it exists alongside this script
SCRIPT_SELF_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
[ -f "${SCRIPT_SELF_DIR}/.env" ] && source "${SCRIPT_SELF_DIR}/.env"

INTERVAL=${INTERVAL:-60}

# ---- Static paths (global — shared by start_worker and run_once) ----
# STACK_DIR: directory containing your akk-stack docker-compose.yml and .env
STACK_DIR="${STACK_DIR:-/opt/eqemu/server}"
# SCRIPT_DIR: directory containing gm_excel.py (mounted as /scripts in worker)
SCRIPT_DIR="${SCRIPT_DIR:-${SCRIPT_SELF_DIR}}"
# EXPORT_DIR: where generated reports are written (mounted as /output in worker)
EXPORT_DIR="${EXPORT_DIR:-/tmp/eqemu-reports}"
# WORKER_NAME: name of the persistent worker Docker container
WORKER_NAME="${WORKER_CONTAINER:-eqemu-gm-worker}"
# MARIADB_CONTAINER: name of the MariaDB Docker container in your akk-stack
MARIADB_CONTAINER="${MARIADB_CONTAINER:-mariadb}"

STATE_FILE="${SCRIPT_DIR}/gm_last_run.txt"

mkdir -p "$EXPORT_DIR" "$SCRIPT_DIR"

# ---- Persistent worker container --------------------------------
# Started once at daemon launch; docker exec replaces docker run --rm.
# This eliminates 8-15s of container cold-start overhead every cycle.
# /tmp is mounted as /data so gm_excel.py's hardcoded /data/ paths
# (last_ts.txt, TSV inputs) resolve correctly without code changes.
start_worker() {
    docker ps --filter "name=${WORKER_NAME}" --filter "status=running" -q \
        | grep -q . && return 0

    echo "[$(date)] Starting persistent report worker..."
    docker rm -f "$WORKER_NAME" 2>/dev/null || true
    docker run -d \
        --name  "$WORKER_NAME" \
        -e TZ="${TZ:-UTC}" \
        -v "/tmp:/data" \
        -v "${SCRIPT_DIR}:/scripts:ro" \
        -v "${EXPORT_DIR}:/output" \
        eqemu-reports \
        tail -f /dev/null
    echo "[$(date)] Worker ready."
}

# ---- Main cycle -------------------------------------------------
run_once() {

TIMESTAMP=$(date +%Y%m%d_%H%M%S)
OUTPUT_XLSX="${EXPORT_DIR}/gm_commands_${TIMESTAMP}.xlsx"

# TSV paths on the host (/tmp) — appear as /data/* inside the container
TSV_CMDS="/tmp/eqemu_gm_commands.tsv"
TSV_KILLS="/tmp/eqemu_gm_kills.tsv"
TSV_ITEMS="/tmp/eqemu_gm_items.tsv"
TSV_IMPACT="/tmp/eqemu_gm_impact.tsv"
TSV_BOTS="/tmp/eqemu_gm_bots.tsv"
TSV_ONLINE="/tmp/eqemu_gm_online.tsv"
TSV_OFFLINE="/tmp/eqemu_gm_offline.tsv"
TSV_CHARS="/tmp/eqemu_gm_chars.tsv"
TSV_DEATHS="/tmp/eqemu_gm_deaths.tsv"
LAST_TS_PREV="/tmp/last_ts.txt"   # gm_excel.py reads this as /data/last_ts.txt

NEW_COUNT=0
LAST_TS=""

# MARIADB_ROOT_PASSWORD can be set directly as an env var, or will be read
# from the akk-stack .env file at STACK_DIR.
if [ -n "$MARIADB_ROOT_PASSWORD" ]; then
  ROOT_PW="$MARIADB_ROOT_PASSWORD"
else
  ROOT_PW=$(grep ^MARIADB_ROOT_PASSWORD "${STACK_DIR}/.env" 2>/dev/null | cut -d= -f2)
fi
if [ -z "$ROOT_PW" ]; then
  echo "ERROR: No MARIADB_ROOT_PASSWORD found. Set it as an env var or ensure"
  echo "       STACK_DIR (${STACK_DIR}) contains a .env with MARIADB_ROOT_PASSWORD."
  exit 1
fi

cd "$STACK_DIR"

# ---- Delta: count commands since last run (serial — must precede queries) ----
if [ -f "$STATE_FILE" ]; then
  LAST_TS=$(cat "$STATE_FILE")
  echo "[$(date)] Last run timestamp: ${LAST_TS}"
  NEW_COUNT=$(docker compose exec -T "${MARIADB_CONTAINER}" mysql -uroot -p"$ROOT_PW" peq \
    --batch --silent --skip-column-names -e "
    SELECT COUNT(*)
    FROM player_event_logs pel
    JOIN character_data cd ON cd.id = pel.character_id
    WHERE pel.event_type_name = 'GM Command'
      AND pel.created_at > '${LAST_TS}';
  " 2>/dev/null)
  echo "[$(date)] New commands since last report: ${NEW_COUNT}"
else
  echo "[$(date)] No previous state file — first run."
fi

# Save previous LAST_TS for gm_excel.py's delta highlighting
echo "$LAST_TS" > "$LAST_TS_PREV"

# ---- Queries — all 8 run concurrently -------------------------
# Each subshell is backgrounded with &; `wait` blocks until all finish.
# Typical saving: reduces ~6s sequential to ~2s (time of slowest query).
echo "[$(date)] Querying (parallel)..."

( docker compose exec -T "${MARIADB_CONTAINER}" mysql -uroot -p"$ROOT_PW" peq \
  --batch --silent -e "
SET time_zone = '${TZ:-UTC}';
SELECT
  DATE_FORMAT(pel.created_at, '%Y-%m-%d %H:%i:%s')        AS timestamp,
  cd.name                                                   AS character_name,
  a.name                                                    AS account_name,
  a.status                                                  AS account_status,
  JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.message')) AS command
FROM player_event_logs pel
JOIN character_data cd ON cd.id = pel.character_id
JOIN account a          ON a.id  = cd.account_id
WHERE pel.event_type_name = 'GM Command'
ORDER BY pel.created_at DESC;
" 2>/dev/null > "$TSV_CMDS" ) &

( docker compose exec -T "${MARIADB_CONTAINER}" mysql -uroot -p"$ROOT_PW" peq \
  --batch --silent -e "
SET time_zone = '${TZ:-UTC}';
SELECT
  DATE_FORMAT(pel.created_at, '%Y-%m-%d %H:%i:%s')                           AS timestamp,
  cd.name                                                                      AS character_name,
  COALESCE(
    JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.target_name')),
    JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.target')),
    'Unknown'
  )                                                                             AS target_name,
  COALESCE(JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.target_id')), '')    AS target_id
FROM player_event_logs pel
JOIN character_data cd ON cd.id = pel.character_id
WHERE pel.event_type_name = 'GM Command'
  AND JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.message')) REGEXP '^#kill'
ORDER BY pel.created_at DESC;
" 2>/dev/null > "$TSV_KILLS" ) &

( docker compose exec -T "${MARIADB_CONTAINER}" mysql -uroot -p"$ROOT_PW" peq \
  --batch --silent -e "
SET time_zone = '${TZ:-UTC}';
SELECT
  DATE_FORMAT(pel.created_at, '%Y-%m-%d %H:%i:%s')                            AS timestamp,
  cd.name                                                                       AS character_name,
  TRIM(SUBSTRING_INDEX(SUBSTRING_INDEX(
    JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.message')), ' ', 2
  ), ' ', -1))                                                                  AS item_id,
  COALESCE(i.Name,      'Unknown Item') AS item_name,
  COALESCE(i.nodrop,    0)              AS nodrop,
  COALESCE(i.reqlevel,  0)             AS reqlevel,
  COALESCE(i.hp,        0)             AS hp,
  COALESCE(i.mana,      0)             AS mana,
  COALESCE(i.ac,        0)             AS ac,
  COALESCE(i.damage,    0)             AS damage,
  COALESCE(i.magic,     0)             AS magic,
  COALESCE(i.lore,      0)             AS lore,
  COALESCE(i.price,     0)             AS price,
  JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.message'))                     AS raw_command
FROM player_event_logs pel
JOIN character_data cd ON cd.id = pel.character_id
LEFT JOIN items i ON i.id = TRIM(SUBSTRING_INDEX(SUBSTRING_INDEX(
    JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.message')), ' ', 2
  ), ' ', -1)) + 0
WHERE pel.event_type_name = 'GM Command'
  AND JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.message'))
      REGEXP '^#(giveitem|summonitem) [0-9]+'
ORDER BY pel.created_at DESC;
" 2>/dev/null > "$TSV_ITEMS" ) &

( docker compose exec -T "${MARIADB_CONTAINER}" mysql -uroot -p"$ROOT_PW" peq \
  --batch --silent -e "
SET time_zone = '${TZ:-UTC}';
SELECT
  DATE_FORMAT(pel.created_at, '%Y-%m-%d %H:%i:%s')                            AS timestamp,
  cd.name                                                                       AS gm_character,
  COALESCE(
    JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.target_name')),
    JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.target')),
    'Unknown'
  )                                                                             AS recipient,
  TRIM(SUBSTRING_INDEX(SUBSTRING_INDEX(
    JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.message')), ' ', 2
  ), ' ', -1))                                                                  AS item_id,
  COALESCE(i.Name,      'Unknown Item') AS item_name,
  COALESCE(i.nodrop,    0)              AS nodrop,
  COALESCE(i.reqlevel,  0)             AS reqlevel,
  COALESCE(i.magic,     0)             AS magic,
  COALESCE(i.lore,      0)             AS lore,
  COALESCE(i.price,     0)             AS price,
  JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.message'))                     AS raw_command
FROM player_event_logs pel
JOIN character_data cd ON cd.id = pel.character_id
LEFT JOIN items i ON i.id = TRIM(SUBSTRING_INDEX(SUBSTRING_INDEX(
    JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.message')), ' ', 2
  ), ' ', -1)) + 0
WHERE pel.event_type_name = 'GM Command'
  AND JSON_UNQUOTE(JSON_EXTRACT(pel.event_data, '\$.message'))
      REGEXP '^#(giveitem|summonitem) [0-9]+'
ORDER BY pel.created_at DESC;
" 2>/dev/null > "$TSV_IMPACT" ) &

( docker compose exec -T "${MARIADB_CONTAINER}" mysql -uroot -p"$ROOT_PW" peq \
  --batch --silent -e "
SET time_zone = '${TZ:-UTC}';
SELECT
  a.name           AS account_name,
  cd.name          AS owner_name,
  b.name           AS bot_name,
  b.class          AS bot_class,
  b.race           AS bot_race,
  b.gender         AS bot_gender,
  b.level          AS bot_level,
  b.hp             AS bot_hp,
  b.mana           AS bot_mana
FROM bot_data b
JOIN character_data cd ON cd.id = b.owner_id
JOIN account a          ON a.id  = cd.account_id
WHERE b.name NOT LIKE '%-deleted-%'
ORDER BY a.name ASC, cd.name ASC, b.name ASC;
" 2>/dev/null > "$TSV_BOTS" ) &

( docker compose exec -T "${MARIADB_CONTAINER}" mysql -uroot -p"$ROOT_PW" peq \
  --batch --silent -e "
SET time_zone = '${TZ:-UTC}';
SELECT cd.name, a.name,
  COALESCE((SELECT z.long_name FROM zone z WHERE z.zoneidnumber=cd.zone_id LIMIT 1),'Unknown'),
  cd.level, cd.class,
  UNIX_TIMESTAMP(NOW()) - cd.last_login AS seconds_online
FROM character_data cd
JOIN account a ON a.id = cd.account_id
WHERE cd.ingame = 1
  AND cd.zone_id > 0
  AND cd.name NOT LIKE '%-deleted-%'
  AND cd.last_login > UNIX_TIMESTAMP(NOW() - INTERVAL 12 HOUR)
ORDER BY cd.name;
" 2>/dev/null > "$TSV_ONLINE" ) &

( docker compose exec -T "${MARIADB_CONTAINER}" mysql -uroot -p"$ROOT_PW" peq \
  --batch --silent -e "
SET time_zone = '${TZ:-UTC}';
SELECT cd.name, a.name,
  cd.level, cd.class,
  DATE_FORMAT(FROM_UNIXTIME(cd.last_login), '%Y-%m-%d %H:%i:%s') AS last_login_time,
  UNIX_TIMESTAMP(NOW()) - cd.last_login AS seconds_offline,
  IF(a.status >= 100, 1, 0) AS is_gm
FROM character_data cd
JOIN account a ON a.id = cd.account_id
WHERE cd.ingame = 0
  AND cd.name NOT LIKE '%-deleted-%'
  AND cd.last_login > UNIX_TIMESTAMP(NOW() - INTERVAL 30 DAY)
ORDER BY cd.last_login DESC
LIMIT 20;
" 2>/dev/null > "$TSV_OFFLINE" ) &

( docker compose exec -T "${MARIADB_CONTAINER}" mysql -uroot -p"$ROOT_PW" peq \
  --batch --silent -e "
SET time_zone = '${TZ:-UTC}';
SELECT
  cd.name,
  cd.class,
  cd.level,
  cd.race,
  a.name        AS account_name,
  a.status      AS account_status,
  COALESCE(g.name,'')  AS guild_name,
  FROM_UNIXTIME(cd.last_login, '%Y-%m-%d %H:%i') AS last_login
FROM character_data cd
JOIN account a ON a.id = cd.account_id
LEFT JOIN guild_members gm ON gm.char_id = cd.id
LEFT JOIN guilds g         ON g.id = gm.guild_id
WHERE cd.name NOT LIKE '%-deleted-%'
ORDER BY cd.last_login DESC;
" 2>/dev/null > "$TSV_CHARS" ) &

# Diagnostic: show raw death event_data on first cycle so field names are visible
if [ ! -f "/tmp/eqemu_death_diag.txt" ]; then
  docker compose exec -T "${MARIADB_CONTAINER}" mysql -uroot -p"$ROOT_PW" peq \
    --batch --silent --skip-column-names -e "
    SELECT event_data FROM player_event_logs
    WHERE event_type_name='Death' LIMIT 1;
  " 2>/dev/null | head -1 > /tmp/eqemu_death_diag.txt
  echo "[$(date)] Death event_data sample: $(cat /tmp/eqemu_death_diag.txt)"
fi

( docker compose exec -T "${MARIADB_CONTAINER}" mysql -uroot -p"$ROOT_PW" peq \
  --batch --silent -e "
SET time_zone = '${TZ:-UTC}';
SELECT
  DATE_FORMAT(pel.created_at, '%Y-%m-%d %H:%i:%s')        AS timestamp,
  cd.name                                                   AS character_name,
  cd.level                                                  AS level,
  COALESCE(
    NULLIF(JSON_UNQUOTE(JSON_EXTRACT(pel.event_data,'\$.killer_name')),     'null'),
    NULLIF(JSON_UNQUOTE(JSON_EXTRACT(pel.event_data,'\$.killed_by_name')), 'null'),
    NULLIF(JSON_UNQUOTE(JSON_EXTRACT(pel.event_data,'\$.killer')),          'null'),
    NULLIF(JSON_UNQUOTE(JSON_EXTRACT(pel.event_data,'\$.killed_by')),       'null'),
    'Unknown'
  )                                                         AS killed_by,
  COALESCE(
    (SELECT z.short_name FROM zone z WHERE z.zoneidnumber = pel.zone_id LIMIT 1),
    NULLIF(JSON_UNQUOTE(JSON_EXTRACT(pel.event_data,'\$.zone_short_name')), 'null'),
    NULLIF(JSON_UNQUOTE(JSON_EXTRACT(pel.event_data,'\$.zone')),            'null'),
    'Unknown'
  )                                                         AS zone,
  COALESCE(
    NULLIF(JSON_UNQUOTE(JSON_EXTRACT(pel.event_data,'\$.spell_id')), 'null'),
    '0'
  )                                                         AS spell_id
FROM player_event_logs pel
JOIN character_data cd ON cd.id = pel.character_id
WHERE pel.event_type_name = 'Death'
ORDER BY pel.created_at DESC;
" 2>/dev/null > "$TSV_DEATHS" ) &

# Block until every background query finishes
wait
echo "[$(date)] Queries done. Commands: $(wc -l < $TSV_CMDS) | Kills: $(wc -l < $TSV_KILLS) | Items: $(wc -l < $TSV_ITEMS) | Impact: $(wc -l < $TSV_IMPACT) | Bots: $(wc -l < $TSV_BOTS)"

# ---- Save latest timestamp for next delta (serial — must follow queries) ----
LATEST_TS=$(docker compose exec -T "${MARIADB_CONTAINER}" mysql -uroot -p"$ROOT_PW" peq \
  --batch --silent --skip-column-names -e "
  SELECT COALESCE(MAX(created_at), NOW())
  FROM player_event_logs
  WHERE event_type_name = 'GM Command';
" 2>/dev/null)
echo "$LATEST_TS" > "$STATE_FILE"
echo "[$(date)] State saved: ${LATEST_TS}"

LINE_COUNT=$(wc -l < "$TSV_CMDS")
if [ "$LINE_COUNT" -le 1 ]; then
  echo "No GM command records found. Skipping build."
  return
fi

# ---- Build report via persistent worker (no container cold-start) ----
echo "[$(date)] Building report..."

# Restart worker if it died between cycles
start_worker

# TSV files are in /tmp on the host; the worker mounts /tmp as /data,
# so all /data/eqemu_gm_*.tsv paths below resolve correctly inside the container.
# Pass --html-only when there are no new commands: skips openpyxl entirely.
# HTML is ALWAYS regenerated so stale content from transient empty TSVs gets
# fixed on the very next cycle even when command counts haven't changed.
EXTRA_FLAGS=""
[ "${NEW_COUNT:-0}" = "0" ] && EXTRA_FLAGS="--html-only"

# Guard: if any key TSV is suspiciously empty, log a warning and skip this
# build cycle so we don't overwrite good HTML with blank sections.
OFFLINE_LINES=$(wc -l < "$TSV_OFFLINE" 2>/dev/null || echo 0)
CHARS_LINES=$(wc -l < "$TSV_CHARS" 2>/dev/null || echo 0)
if [ "$OFFLINE_LINES" -lt 1 ] && [ "$CHARS_LINES" -lt 1 ]; then
  echo "[$(date)] WARNING: offline and chars TSVs both empty — skipping build to preserve previous index.html"
  return
fi

docker exec "$WORKER_NAME" \
  python /scripts/gm_excel.py \
    /data/eqemu_gm_commands.tsv \
    /data/eqemu_gm_kills.tsv \
    /data/eqemu_gm_items.tsv \
    /data/eqemu_gm_impact.tsv \
    /data/eqemu_gm_bots.tsv \
    /output/gm_commands_${TIMESTAMP}.xlsx \
    /output/index.html \
    ${NEW_COUNT} \
    /data/eqemu_gm_chars.tsv \
    /data/eqemu_gm_deaths.tsv \
    ${EXTRA_FLAGS}

# ---- Keep only the single latest xlsx ----
ls -t "${EXPORT_DIR}"/gm_commands_*.xlsx 2>/dev/null | tail -n +2 | xargs -r rm --

if [ -f "${EXPORT_DIR}/index.html" ]; then
  if [ -n "$EXTRA_FLAGS" ]; then
    echo "[$(date)] Excel: skipped (no new commands — existing xlsx unchanged)"
  else
    echo "[$(date)] Excel: ${OUTPUT_XLSX}"
  fi
  echo "[$(date)] HTML:  index.html updated"
  echo "[$(date)] Web:   http://localhost:${DASHBOARD_PORT:-8765}"
else
  echo "[$(date)] ERROR: index.html was not created."
fi

echo "[$(date)] Done."
} # end run_once()

# ---- Daemon loop ------------------------------------------------
echo "[$(date)] GM report daemon started — refreshing every ${INTERVAL}s"
start_worker
while true; do
    run_once
    sleep "$INTERVAL"
done
