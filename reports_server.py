# ==============================================================
# EQEmu GM Report — HTTP Server
# Project: https://github.com/YOUR_USERNAME/eqemu-gm-dashboard
# Version: 1.3  |  Created: May 2026
#
# Lightweight threaded HTTP server that serves the GM Command
# Usage report, handles server-side clear_time.json tracking,
# proxies Spire API requests, and serves character/bot
# inventory data directly from MariaDB via PyMySQL.
#
# Free to use and adapt. Please keep this attribution intact.
# ==============================================================

import os, json, threading, urllib.request, urllib.error, urllib.parse
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn

# ── Configuration from environment variables ─────────────────────────────────
# Set these in your .env file (see .env.example).
REPORTS_DIR      = os.environ.get("REPORTS_DIR",    "/reports")
SPIRE_BASE_URL   = os.environ.get("SPIRE_URL",      "http://localhost:3000")
PORT             = int(os.environ.get("DASHBOARD_PORT", "8765"))

# DB connection — env vars take priority, db_config.json used as fallback
_DB_HOST         = os.environ.get("DB_HOST", "")
_DB_PORT         = int(os.environ.get("DB_PORT", "3306"))
_DB_USER         = os.environ.get("DB_USER", "root")
_DB_PASS         = os.environ.get("DB_PASS", "")
_DB_NAME         = os.environ.get("DB_NAME", "peq")

CLEAR_FILE       = os.path.join(REPORTS_DIR, "clear_time.json")
SPIRE_TOKEN_FILE = os.path.join(REPORTS_DIR, "spire_token.txt")
DB_CONFIG_FILE   = os.path.join(REPORTS_DIR, "db_config.json")

EQ_SLOTS = {
    0:"Charm", 1:"Ear", 2:"Head", 3:"Face", 4:"Ear",
    5:"Neck", 6:"Shoulder", 7:"Arms", 8:"Back",
    9:"L.Wrist", 10:"R.Wrist", 11:"Hands",
    12:"Primary", 13:"Secondary",
    14:"L.Ring", 15:"R.Ring",
    16:"Chest", 17:"Legs", 18:"Feet", 19:"Waist",
    20:"P.Source", 21:"Ammo"
}

# ── DB helpers ────────────────────────────────────────────────

def _db_connect():
    """Connect to MariaDB/MySQL.
    Priority: environment variables > db_config.json > built-in defaults."""
    try:
        import pymysql
        host, port, user, password, database = (
            _DB_HOST, _DB_PORT, _DB_USER, _DB_PASS, _DB_NAME
        )
        # Fall back to db_config.json if env vars not set
        if not host:
            try:
                with open(DB_CONFIG_FILE) as f:
                    cfg = json.load(f)
                host     = cfg.get("host",     "127.0.0.1")
                port     = int(cfg.get("port", 3306))
                user     = cfg.get("user",     "root")
                password = cfg.get("password", "")
                database = cfg.get("db",       "peq")
            except Exception:
                return None
        return pymysql.connect(
            host=host, port=port, user=user,
            password=password, database=database,
            charset="utf8mb4",
            cursorclass=pymysql.cursors.DictCursor
        )
    except Exception:
        return None

def _fetch_inventory(name, is_bot=False):
    conn = _db_connect()
    if not conn:
        return None, "DB not configured — set DB_HOST/DB_PASS env vars or create db_config.json in REPORTS_DIR"
    try:
        with conn.cursor() as cur:
            if is_bot:
                cur.execute(
                    "SELECT bot_id AS id, name, class, level, race, gender "
                    "FROM bot_data WHERE name=%s LIMIT 1", (name,)
                )
            else:
                cur.execute(
                    "SELECT id, name, class, level, race, gender "
                    "FROM character_data WHERE name=%s LIMIT 1", (name,)
                )
            char = cur.fetchone()
            if not char:
                return None, f"{'Bot' if is_bot else 'Character'} '{name}' not found"

            char_id = char["id"]
            if is_bot:
                cur.execute("""
                    SELECT bi.slot_id AS slotid, bi.item_id AS itemid,
                           COALESCE(i.Name,'') AS item_name,
                           COALESCE(i.ac,0) ac, COALESCE(i.hp,0) hp, COALESCE(i.mana,0) mana,
                           COALESCE(i.magic,0) magic, COALESCE(i.nodrop,0) nodrop,
                           COALESCE(i.norent,0) norent,
                           COALESCE(i.astr,0) astr, COALESCE(i.asta,0) asta,
                           COALESCE(i.aagi,0) aagi, COALESCE(i.adex,0) adex,
                           COALESCE(i.aint,0) aint, COALESCE(i.awis,0) awis,
                           COALESCE(i.acha,0) acha,
                           COALESCE(i.damage,0) damage, COALESCE(i.delay,0) delay,
                           COALESCE(i.weight,0) weight,
                           COALESCE(i.reqlevel,0) reqlevel, COALESCE(i.id,0) item_id_real
                    FROM bot_inventories bi
                    LEFT JOIN items i ON i.id=bi.item_id
                    WHERE bi.bot_id=%s AND bi.slot_id BETWEEN 0 AND 21
                    ORDER BY bi.slot_id
                """, (char_id,))
            else:
                cur.execute("""
                    SELECT inv.slot_id AS slotid, inv.item_id AS itemid,
                           COALESCE(i.Name,'') AS item_name,
                           COALESCE(i.ac,0) ac, COALESCE(i.hp,0) hp, COALESCE(i.mana,0) mana,
                           COALESCE(i.magic,0) magic, COALESCE(i.nodrop,0) nodrop,
                           COALESCE(i.norent,0) norent,
                           COALESCE(i.astr,0) astr, COALESCE(i.asta,0) asta,
                           COALESCE(i.aagi,0) aagi, COALESCE(i.adex,0) adex,
                           COALESCE(i.aint,0) aint, COALESCE(i.awis,0) awis,
                           COALESCE(i.acha,0) acha,
                           COALESCE(i.damage,0) damage, COALESCE(i.delay,0) delay,
                           COALESCE(i.weight,0) weight,
                           COALESCE(i.reqlevel,0) reqlevel, COALESCE(i.id,0) item_id_real
                    FROM inventory inv
                    LEFT JOIN items i ON i.id=inv.item_id
                    WHERE inv.character_id=%s AND inv.slot_id BETWEEN 0 AND 21
                    ORDER BY inv.slot_id
                """, (char_id,))

            slots = cur.fetchall()
        return {"character": char, "slots": slots, "is_bot": is_bot}, None
    except Exception as e:
        return None, str(e)
    finally:
        conn.close()

# ── HTTP server ───────────────────────────────────────────────

class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True

class ReportHandler(SimpleHTTPRequestHandler):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=REPORTS_DIR, **kwargs)

    def do_GET(self):
        self._full_path = self.path          # preserve query string for param handlers
        self.path = self.path.split('?')[0]
        if self.path.startswith("/spire/"):
            self._proxy_spire()
        elif self.path == "/last-modified":
            self._handle_last_modified()
        elif self.path == "/levels":
            self._handle_levels()
        elif self.path.startswith("/spell-search"):
            self._handle_spell_search()
        elif self.path == "/spire-user-log":
            self._handle_spire_user_log()
        elif self.path.startswith("/inventory/"):
            self._handle_inventory()
        else:
            super().do_GET()

    def do_POST(self):
        if self.path == "/clear":
            ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            data = json.dumps({"cleared_at": ts}).encode()
            with open(CLEAR_FILE, "w") as f:
                f.write(data.decode())
            self._respond(200, data)
        else:
            self._respond(404, b'{"error":"not found"}')

    def do_OPTIONS(self):
        self._respond(200, b"")

    def _proxy_spire(self):
        spire_path = self.path[7:]
        url        = f"{SPIRE_BASE_URL}/{spire_path}"
        token      = ""
        try:
            with open(SPIRE_TOKEN_FILE) as f:
                token = f.read().strip()
        except Exception:
            pass
        try:
            req  = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
            resp = urllib.request.urlopen(req, timeout=5)
            data = resp.read()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(data)
        except urllib.error.HTTPError as e:
            self._respond(e.code, e.read())
        except Exception as e:
            self._respond(503, json.dumps({"error": str(e)}).encode())

    def _handle_last_modified(self):
        try:
            mtime = int(os.path.getmtime(os.path.join(REPORTS_DIR, "index.html")))
        except Exception:
            mtime = 0
        body = json.dumps({"mtime": mtime}).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _handle_levels(self):
        conn = _db_connect()
        if not conn:
            self._respond(503, b'{"error":"DB not configured"}')
            return
        try:
            with conn.cursor() as cur:
                cur.execute("""
                    SELECT cd.name,
                           pel.created_at AS ts_raw,
                           DATE_FORMAT(pel.created_at, '%Y-%m-%d %H:%i') AS ts,
                           CAST(JSON_UNQUOTE(JSON_EXTRACT(pel.event_data,'$.from_level')) AS UNSIGNED) AS from_level,
                           CAST(JSON_UNQUOTE(JSON_EXTRACT(pel.event_data,'$.to_level'))   AS UNSIGNED) AS to_level
                    FROM player_event_logs pel
                    JOIN character_data cd ON cd.id = pel.character_id
                    WHERE pel.event_type_name IN ('Level Gain','Level Loss')
                    ORDER BY cd.name, pel.created_at
                """)
                rows = cur.fetchall()
            from datetime import datetime
            chars = {}
            for r in rows:
                n = r["name"]
                if n not in chars:
                    chars[n] = []
                prev = chars[n][-1] if chars[n] else None
                if prev and r["ts_raw"] and prev.get("ts_raw"):
                    try:
                        t1 = prev["ts_raw"] if isinstance(prev["ts_raw"], datetime) else datetime.strptime(str(prev["ts_raw"]), "%Y-%m-%d %H:%M:%S")
                        t2 = r["ts_raw"] if isinstance(r["ts_raw"], datetime) else datetime.strptime(str(r["ts_raw"]), "%Y-%m-%d %H:%M:%S")
                        elapsed = round((t2 - t1).total_seconds() / 60, 1)
                    except:
                        elapsed = None
                else:
                    elapsed = None
                jumped = int(r["to_level"]) - int(r["from_level"])
                chars[n].append({
                    "ts": r["ts"],
                    "ts_raw": str(r["ts_raw"]),
                    "from_level": int(r["from_level"]),
                    "to_level": int(r["to_level"]),
                    "jumped": jumped,
                    "elapsed_min": elapsed
                })
            # Strip ts_raw from output
            for n in chars:
                for e in chars[n]:
                    e.pop("ts_raw", None)
            body = json.dumps(chars, default=str).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(body)
        except Exception as e:
            self._respond(503, json.dumps({"error": str(e)}).encode())
        finally:
            conn.close()

    def _handle_spell_search(self):
        """Search spells_new by name or ID.
        GET /spell-search?name=Bliss   → up to 5 matches
        GET /spell-search?id=12345     → exact ID lookup
        """
        import urllib.parse as _up
        qs     = _up.parse_qs(_up.urlparse(getattr(self,'_full_path',self.path)).query)
        name   = qs.get("name",  [None])[0]
        sid    = qs.get("id",    [None])[0]

        conn = _db_connect()
        if not conn:
            self._respond(503, b'{"error":"DB not configured"}')
            return
        try:
            with conn.cursor() as cur:
                if sid:
                    cur.execute("SELECT * FROM spells_new WHERE id=%s LIMIT 1",
                                (int(sid),))
                elif name:
                    # Build a short prefix from the first 3 words for fallback matching
                    words = name.split()
                    prefix = " ".join(words[:3]) if len(words) > 3 else name
                    # Strip fancy apostrophes → plain apostrophe for comparison
                    name_plain   = name.replace("\u2019","'").replace("\u2018","'")
                    prefix_plain = prefix.replace("\u2019","'").replace("\u2018","'")
                    cur.execute(
                        "SELECT * FROM spells_new WHERE "
                        "name=%s OR name=%s OR "
                        "name LIKE %s OR name LIKE %s OR "
                        "name LIKE %s "
                        "LIMIT 10",
                        (name, name_plain,
                         f"%{name_plain}%",
                         f"{prefix_plain}%",
                         f"{prefix}%")
                    )
                else:
                    self._respond(400, b'{"error":"supply name= or id="}')
                    return
                rows = cur.fetchall()
            body = json.dumps(rows, default=str).encode()
            self.send_response(200)
            self.send_header("Content-Type",  "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(body)
        except Exception as e:
            self._respond(503, json.dumps({"error": str(e)}).encode())
        finally:
            conn.close()


    def _handle_spire_user_log(self):
        conn = _db_connect()
        if not conn:
            self._respond(503, b'{"error":"DB not configured"}')
            return
        try:
            with conn.cursor() as cur:
                cur.execute("""
                    SELECT
                        suel.id,
                        DATE_FORMAT(suel.created_at, '%Y-%m-%d %H:%i:%s') AS created_at,
                        COALESCE(su.user_name, CONCAT('User #', suel.user_id)) AS user_name,
                        suel.user_id,
                        suel.event_name,
                        COALESCE(suel.data, '')                              AS data
                    FROM spire_user_event_log suel
                    LEFT JOIN spire_users su ON su.id = suel.user_id
                    ORDER BY suel.created_at DESC
                    LIMIT 2000
                """)
                rows = cur.fetchall()
            body = json.dumps(rows, default=str).encode()
            self.send_response(200)
            self.send_header("Content-Type",  "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(body)
        except Exception as e:
            self._respond(503, json.dumps({"error": str(e)}).encode())
        finally:
            conn.close()


    def _handle_inventory(self):
        parts  = self.path.split("/")           # ['','inventory','character','Powah']
        if len(parts) < 4:
            self._respond(400, b'{"error":"usage: /inventory/character/:name or /inventory/bot/:name"}')
            return
        inv_type = parts[2]
        name     = urllib.parse.unquote(parts[3])
        if inv_type not in ("character", "bot"):
            self._respond(400, b'{"error":"type must be character or bot"}')
            return
        data, err = _fetch_inventory(name, is_bot=(inv_type == "bot"))
        if err:
            self._respond(404, json.dumps({"error": err}).encode())
        else:
            body = json.dumps(data, default=str).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(body)

    def _respond(self, code, body):
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass

if __name__ == "__main__":
    os.makedirs(REPORTS_DIR, exist_ok=True)
    if not os.path.exists(CLEAR_FILE):
        with open(CLEAR_FILE, "w") as f:
            json.dump({"cleared_at": ""}, f)

    server = ThreadedHTTPServer(("0.0.0.0", PORT), ReportHandler)
    print(f"EQEmu GM Dashboard running on http://0.0.0.0:{PORT}")
    server.serve_forever()
