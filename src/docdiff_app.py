#!/usr/bin/env python3
"""
DocDiff — local web UI for document comparison.

Run:  python docdiff_app.py            (opens browser at http://127.0.0.1:8765)
      python docdiff_app.py 9000       (custom port)

Drop two documents (.pdf / .docx / .txt) onto the page — or click to browse —
and get a side-by-side word-level diff. Reports are also saved to ./reports.
"""

import argparse
import email
import json
import os
import socket
import sys
import tempfile
import threading
import uuid
import webbrowser
from email import policy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import compare_docs as engine

DEFAULT_PORT = 8765
ALLOWED_EXT = {".pdf", ".docx", ".txt"}
MAX_UPLOAD = 200 * 1024 * 1024  # 200 MB combined

JOBS = {}          # job_id -> {status, phase, report, summary, error, saved}
JOBS_LOCK = threading.Lock()


def application_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


REPORTS_DIR = application_dir() / "reports"


# ══════════════════════════════════════════════════════════════
#  UPLOAD PAGE — light glassmorphism, pastel, amber accents
# ══════════════════════════════════════════════════════════════

INDEX_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DocDiff — Compare documents</title>
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --ink:        #2b2a33;
  --ink-soft:   #6d6f7e;
  --ink-faint:  #9b9eae;
  --amber:      #d97706;
  --amber-deep: #b45309;
  --glass:      rgba(255,255,255,0.62);
  --glass-edge: rgba(255,255,255,0.85);
}

html, body { height: 100%; }

body {
  font-family: 'Inter', system-ui, sans-serif;
  min-height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 28px;
  color: var(--ink);
  background: linear-gradient(165deg, #fdfbf7 0%, #f3f5fb 48%, #faf3ea 100%);
  overflow-x: hidden;
}

/* drifting pastel blobs */
.blob {
  position: fixed;
  border-radius: 50%;
  filter: blur(70px);
  opacity: .55;
  z-index: 0;
  pointer-events: none;
  animation: drift 26s ease-in-out infinite alternate;
}
.blob.b1 { width: 480px; height: 480px; left: -140px; top: -120px;
           background: radial-gradient(circle, #ffe7b8, transparent 70%); }
.blob.b2 { width: 420px; height: 420px; right: -120px; top: 8%;
           background: radial-gradient(circle, #cfe6ff, transparent 70%);
           animation-delay: -8s; animation-duration: 32s; }
.blob.b3 { width: 520px; height: 520px; left: 30%; bottom: -220px;
           background: radial-gradient(circle, #e8ddff, transparent 70%);
           animation-delay: -16s; animation-duration: 38s; }

@keyframes drift {
  0%   { transform: translate(0, 0) scale(1); }
  100% { transform: translate(60px, 40px) scale(1.12); }
}

.wrap { width: 100%; max-width: 720px; position: relative; z-index: 1; }

/* entrance */
@keyframes rise { from { opacity: 0; transform: translateY(18px); }
                  to   { opacity: 1; transform: none; } }
.hero  { animation: rise .55s cubic-bezier(.2,.7,.3,1) both; }
.card  { animation: rise .55s cubic-bezier(.2,.7,.3,1) .12s both; }

.hero { text-align: center; margin-bottom: 26px; }
.hero .logo {
  width: 56px; height: 56px; margin: 0 auto 14px;
  border-radius: 16px;
  display: flex; align-items: center; justify-content: center;
  font-size: 25px; font-weight: 800; color: #fff;
  background: linear-gradient(135deg, #fbbf24, #d97706);
  box-shadow: 0 10px 28px rgba(217,119,6,0.30);
  animation: bob 5s ease-in-out infinite;
}
@keyframes bob { 0%,100% { transform: translateY(0); } 50% { transform: translateY(-5px); } }
.hero h1 { font-size: 27px; font-weight: 800; letter-spacing: -0.5px; }
.hero p  { margin-top: 6px; font-size: 13.5px; color: var(--ink-soft); }

/* frosted workspace card */
.card {
  background: var(--glass);
  border: 1px solid var(--glass-edge);
  border-radius: 24px;
  backdrop-filter: blur(26px) saturate(160%);
  -webkit-backdrop-filter: blur(26px) saturate(160%);
  box-shadow: 0 20px 60px rgba(80, 70, 50, 0.13), inset 0 1px 0 rgba(255,255,255,0.9);
  padding: 26px;
}

/* ── single dropzone ── */
.drop {
  border: 1.5px dashed rgba(217,119,6,0.40);
  border-radius: 18px;
  background: rgba(255,255,255,0.45);
  padding: 38px 22px;
  text-align: center;
  cursor: pointer;
  transition: border-color .2s, background .2s, transform .2s, box-shadow .2s;
}
.drop:hover {
  background: rgba(255,251,240,0.75);
  border-color: var(--amber);
  transform: translateY(-2px);
  box-shadow: 0 12px 30px rgba(217,119,6,0.12);
}
.drop.dragover {
  background: rgba(255,247,228,0.9);
  border-color: var(--amber);
  border-style: solid;
  transform: scale(1.012);
  box-shadow: 0 14px 36px rgba(217,119,6,0.18);
}
.drop .di { font-size: 30px; display: block; margin-bottom: 10px; }
.drop .dt { font-size: 15px; font-weight: 700; }
.drop .dt span { color: var(--amber-deep); }
.drop .dh { margin-top: 5px; font-size: 12.5px; color: var(--ink-soft); }
.drop input { display: none; }

/* ── slots ── */
.slots {
  display: grid;
  grid-template-columns: 1fr 40px 1fr;
  gap: 10px;
  align-items: center;
  margin-top: 14px;
}
@media (max-width: 560px) { .slots { grid-template-columns: 1fr; }
  .swap { transform: rotate(90deg); margin: 0 auto; } }

.slot {
  border-radius: 14px;
  border: 1px solid rgba(43,42,51,0.08);
  background: rgba(255,255,255,0.55);
  padding: 11px 14px;
  min-height: 64px;
  display: flex; flex-direction: column; justify-content: center; gap: 3px;
  position: relative;
  transition: background .2s, border-color .2s;
}
.slot.filled {
  background: rgba(255,255,255,0.85);
  border-color: rgba(217,119,6,0.25);
  animation: pop .3s cubic-bezier(.2,.8,.3,1.2);
}
@keyframes pop { from { transform: scale(.96); opacity: .4; } to { transform: scale(1); opacity: 1; } }

.slot .tag {
  font-size: 9.5px; font-weight: 800; letter-spacing: 1.1px; text-transform: uppercase;
}
.slot.base .tag { color: #be5a3f; }
.slot.rev  .tag { color: #2f8a5d; }
.slot .fname {
  font-size: 12.5px; font-weight: 600; color: var(--ink);
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 92%;
}
.slot .empty-hint { font-size: 12px; color: var(--ink-faint); }
.slot .clear {
  position: absolute; top: 8px; right: 9px;
  border: none; background: rgba(43,42,51,0.06); color: var(--ink-soft);
  width: 20px; height: 20px; border-radius: 50%;
  font-size: 11px; line-height: 1; cursor: pointer; display: none;
  transition: background .15s, color .15s;
}
.slot.filled .clear { display: block; }
.slot .clear:hover { background: rgba(190,90,63,0.15); color: #be5a3f; }

.swap {
  width: 38px; height: 38px; border-radius: 50%;
  border: 1px solid rgba(43,42,51,0.10);
  background: rgba(255,255,255,0.7);
  color: var(--ink-soft); font-size: 15px; cursor: pointer;
  transition: transform .25s, color .15s, border-color .15s;
}
.swap:hover { transform: rotate(180deg); color: var(--amber-deep); border-color: rgba(217,119,6,0.35); }

/* ── compare button ── */
.go {
  width: 100%;
  margin-top: 18px;
  border: none; border-radius: 15px;
  padding: 15px;
  font-family: inherit; font-size: 15px; font-weight: 700; letter-spacing: .2px;
  color: #fff;
  background: linear-gradient(135deg, #f59e0b, #d97706);
  cursor: pointer;
  position: relative; overflow: hidden;
  transition: transform .18s, box-shadow .18s, opacity .18s;
  box-shadow: 0 10px 26px rgba(217,119,6,0.28);
}
.go::after {            /* sheen sweep */
  content: '';
  position: absolute; top: 0; left: -70%;
  width: 50%; height: 100%;
  background: linear-gradient(105deg, transparent, rgba(255,255,255,0.45), transparent);
  transform: skewX(-20deg);
  transition: left .5s ease;
}
.go:hover:not(:disabled)::after { left: 130%; }
.go:hover:not(:disabled) { transform: translateY(-2px); box-shadow: 0 14px 34px rgba(217,119,6,0.38); }
.go:disabled { opacity: .35; cursor: not-allowed; box-shadow: none; }

.note { margin-top: 14px; text-align: center; font-size: 11.5px; color: var(--ink-faint); }

/* ── progress ── */
.progress { display: none; margin-top: 20px; text-align: center; animation: rise .35s both; }
.progress.active { display: block; }
.bar {
  height: 6px; border-radius: 4px; overflow: hidden;
  background: rgba(217,119,6,0.12); margin-bottom: 12px;
}
.bar i {
  display: block; height: 100%; width: 38%;
  border-radius: 4px;
  background: linear-gradient(90deg, #fbbf24, #d97706);
  animation: slide 1.25s ease-in-out infinite;
}
@keyframes slide { 0% { margin-left: -40%; } 100% { margin-left: 105%; } }
.phase { font-size: 13px; font-weight: 600; color: var(--ink); }
.elapsed { font-size: 11px; color: var(--ink-faint); margin-top: 4px; }

/* ── error ── */
.error {
  display: none; margin-top: 16px; padding: 12px 16px;
  border-radius: 13px; font-size: 13px;
  background: rgba(190,90,63,0.10); border: 1px solid rgba(190,90,63,0.30); color: #a3492f;
  animation: rise .3s both;
}
.error.active { display: block; }
</style>
</head>
<body>
<div class="blob b1"></div>
<div class="blob b2"></div>
<div class="blob b3"></div>

<div class="wrap">
  <div class="hero">
    <div class="logo">&#8644;</div>
    <h1>DocDiff</h1>
    <p>Side-by-side comparison with word-level changes &mdash; PDF, Word &amp; text</p>
  </div>

  <div class="card">
    <div class="drop" id="drop">
      <span class="di">&#128194;</span>
      <span class="dt">Drop two documents here &mdash; or <span>click to browse</span></span>
      <div class="dh">First file = original &middot; second = revised. You can also add them one at a time.</div>
      <input type="file" id="fileInput" accept=".pdf,.docx,.txt" multiple>
    </div>

    <div class="slots">
      <div class="slot base" id="slotA">
        <button class="clear" type="button" title="Remove">&#10005;</button>
        <span class="tag">Base &middot; original</span>
        <span class="fname"></span>
        <span class="empty-hint">No file yet</span>
      </div>
      <button class="swap" id="swapBtn" type="button" title="Swap documents">&#8644;</button>
      <div class="slot rev" id="slotB">
        <button class="clear" type="button" title="Remove">&#10005;</button>
        <span class="tag">Revised &middot; newer</span>
        <span class="fname"></span>
        <span class="empty-hint">No file yet</span>
      </div>
    </div>

    <button class="go" id="goBtn" type="button" disabled>Compare documents</button>

    <div class="progress" id="progress">
      <div class="bar"><i></i></div>
      <div class="phase" id="phase">Uploading&hellip;</div>
      <div class="elapsed" id="elapsed"></div>
    </div>

    <div class="error" id="error"></div>

    <div class="note">Runs entirely on this computer &mdash; your files never leave it.</div>
  </div>
</div>

<script>
'use strict';
var state = { a: null, b: null, busy: false };

var drop      = document.getElementById('drop');
var fileInput = document.getElementById('fileInput');
var slotA     = document.getElementById('slotA');
var slotB     = document.getElementById('slotB');
var goBtn     = document.getElementById('goBtn');
var progress  = document.getElementById('progress');
var phaseEl   = document.getElementById('phase');
var elapsedEl = document.getElementById('elapsed');
var errEl     = document.getElementById('error');

function isOk(f) { return /\\.(pdf|docx|txt)$/i.test(f.name); }

function showError(msg) { errEl.textContent = msg; errEl.classList.add('active'); }
function hideError()    { errEl.classList.remove('active'); }

function renderSlot(slot, file) {
  var fname = slot.querySelector('.fname');
  var hint  = slot.querySelector('.empty-hint');
  if (file) {
    slot.classList.add('filled');
    fname.textContent = file.name;
    hint.style.display = 'none';
  } else {
    slot.classList.remove('filled');
    fname.textContent = '';
    hint.style.display = '';
  }
}

function refresh() {
  renderSlot(slotA, state.a);
  renderSlot(slotB, state.b);
  goBtn.disabled = !(state.a && state.b) || state.busy;
}

function addFiles(list) {
  hideError();
  var files = [];
  for (var i = 0; i < list.length; i++) files.push(list[i]);
  var bad = files.filter(function (f) { return !isOk(f); });
  if (bad.length) {
    showError('Unsupported file type: ' + bad[0].name + ' \\u2014 use .pdf, .docx or .txt');
  }
  files = files.filter(isOk).slice(0, 2);
  if (!files.length) return;

  if (files.length === 2) {              // pair dropped: base first, revised second
    state.a = files[0];
    state.b = files[1];
  } else {                               // single file: fill the first empty slot
    if (!state.a)      state.a = files[0];
    else if (!state.b) state.b = files[0];
    else               state.b = files[0];   // both full: replace revised
  }
  refresh();
}

/* click to browse — explicit, works everywhere */
drop.addEventListener('click', function () { fileInput.click(); });
fileInput.addEventListener('change', function () {
  addFiles(fileInput.files);
  fileInput.value = '';
});

/* drag & drop on the zone and the whole page */
['dragover', 'dragenter'].forEach(function (ev) {
  document.addEventListener(ev, function (e) {
    e.preventDefault();
    drop.classList.add('dragover');
  });
});
['dragleave', 'dragend'].forEach(function (ev) {
  document.addEventListener(ev, function (e) {
    if (ev === 'dragleave' && e.relatedTarget) return;
    drop.classList.remove('dragover');
  });
});
document.addEventListener('drop', function (e) {
  e.preventDefault();
  drop.classList.remove('dragover');
  if (e.dataTransfer && e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
});

/* slot clear buttons */
slotA.querySelector('.clear').addEventListener('click', function (e) {
  e.stopPropagation(); state.a = null; refresh();
});
slotB.querySelector('.clear').addEventListener('click', function (e) {
  e.stopPropagation(); state.b = null; refresh();
});

/* swap */
document.getElementById('swapBtn').addEventListener('click', function () {
  var t = state.a; state.a = state.b; state.b = t;
  refresh();
});

/* compare */
goBtn.addEventListener('click', function () {
  if (!state.a || !state.b || state.busy) return;
  state.busy = true;
  refresh();
  hideError();
  progress.classList.add('active');
  phaseEl.textContent = 'Uploading\\u2026';
  var t0 = Date.now();
  var timer = setInterval(function () {
    elapsedEl.textContent = Math.round((Date.now() - t0) / 1000) + 's elapsed';
  }, 1000);

  function fail(msg) {
    clearInterval(timer);
    progress.classList.remove('active');
    state.busy = false;
    refresh();
    showError(msg);
  }

  var fd = new FormData();
  fd.append('base', state.a);
  fd.append('revised', state.b);

  fetch('/compare', { method: 'POST', body: fd })
    .then(function (r) { return r.json().then(function (j) { return { ok: r.ok, j: j }; }); })
    .then(function (res) {
      if (!res.ok) return fail(res.j.error || 'Upload failed');
      var jid = res.j.job;
      var poll = setInterval(function () {
        fetch('/status/' + jid)
          .then(function (r) { return r.json(); })
          .then(function (s) {
            if (s.status === 'done') {
              clearInterval(poll); clearInterval(timer);
              phaseEl.textContent = 'Done \\u2014 opening report\\u2026';
              window.location.href = '/report/' + jid;
            } else if (s.status === 'error') {
              clearInterval(poll);
              fail(s.error || 'Comparison failed');
            } else if (s.phase) {
              phaseEl.textContent = s.phase;
            }
          })
          .catch(function () { /* transient poll failure — keep polling */ });
      }, 700);
    })
    .catch(function (e) { fail('Could not reach the local server: ' + e.message); });
});
</script>
</body>
</html>
"""


# ══════════════════════════════════════════════════════════════
#  COMPARISON JOB
# ══════════════════════════════════════════════════════════════

def run_job(job_id, path_a, path_b, name_a, name_b):
    job = JOBS[job_id]
    try:
        job["phase"] = f"Reading {name_a} …"
        blocks_a = engine.extract_blocks(path_a)
        job["phase"] = f"Reading {name_b} …"
        blocks_b = engine.extract_blocks(path_b)
        job["phase"] = "Aligning content and computing differences …"
        rows = engine.align_blocks(blocks_a, blocks_b)
        summary = engine.summarise(rows)
        job["phase"] = "Rendering report …"
        report = engine.build_html(rows, summary, name_a, name_b, home_link=True)
        job["report"] = report
        job["summary"] = summary
        # Keep a copy on disk so the user can share/archive it
        try:
            outdir = REPORTS_DIR
            outdir.mkdir(exist_ok=True)
            safe = lambda s: "".join(c if c.isalnum() or c in "-_." else "_" for c in s)
            fn = outdir / f"diff_{safe(Path(name_a).stem)[:60]}_vs_{safe(Path(name_b).stem)[:60]}.html"
            fn.write_text(report, encoding="utf-8")
            job["saved"] = str(fn)
        except OSError:
            pass
        job["status"] = "done"
    except SystemExit as e:          # engine signals unsupported input this way
        job.update(status="error", error=str(e))
    except Exception as e:
        job.update(status="error", error=f"{type(e).__name__}: {e}")
    finally:
        for p in (path_a, path_b):
            try:
                os.unlink(p)
            except OSError:
                pass


# ══════════════════════════════════════════════════════════════
#  HTTP SERVER
# ══════════════════════════════════════════════════════════════

def parse_multipart(content_type, body):
    msg = email.message_from_bytes(
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode()
        + body, policy=policy.default)
    parts = {}
    for part in msg.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if name:
            parts[name] = (part.get_filename() or "", part.get_payload(decode=True) or b"")
    return parts


class Handler(BaseHTTPRequestHandler):
    server_version = "DocDiff/1.1"

    def log_message(self, *args):   # keep the console quiet
        pass

    def _send(self, code, body, ctype="text/html; charset=utf-8"):
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def _json(self, code, obj):
        self._send(code, json.dumps(obj), "application/json")

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            return self._send(200, INDEX_HTML)
        if self.path.startswith("/status/"):
            job = JOBS.get(self.path.rsplit("/", 1)[-1])
            if not job:
                return self._json(404, {"error": "Unknown job"})
            return self._json(200, {k: job.get(k) for k in
                                    ("status", "phase", "error", "summary", "saved")})
        if self.path.startswith("/report/"):
            job = JOBS.get(self.path.rsplit("/", 1)[-1])
            if not job or job.get("status") != "done":
                return self._send(404, "<h3>Report not found.</h3><a href='/'>Back</a>")
            return self._send(200, job["report"])
        return self._send(404, "<h3>Not found.</h3><a href='/'>Back</a>")

    def do_POST(self):
        if self.path != "/compare":
            return self._json(404, {"error": "Unknown endpoint"})
        try:
            length = int(self.headers.get("Content-Length", 0))
            if length <= 0:
                return self._json(400, {"error": "Empty request"})
            if length > MAX_UPLOAD:
                return self._json(413, {"error": "Files too large (200 MB limit)"})
            ctype = self.headers.get("Content-Type", "")
            if "multipart/form-data" not in ctype:
                return self._json(400, {"error": "Expected multipart form data"})

            parts = parse_multipart(ctype, self.rfile.read(length))
            if "base" not in parts or "revised" not in parts:
                return self._json(400, {"error": "Both documents are required"})

            paths, names = [], []
            for key in ("base", "revised"):
                fname, data = parts[key]
                ext = Path(fname).suffix.lower()
                if ext not in ALLOWED_EXT:
                    return self._json(400, {"error": f"Unsupported file type: {fname}"})
                if not data:
                    return self._json(400, {"error": f"{fname} is empty"})
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
                tmp.write(data)
                tmp.close()
                paths.append(tmp.name)
                names.append(fname)

            job_id = uuid.uuid4().hex[:12]
            with JOBS_LOCK:
                JOBS[job_id] = {"status": "running", "phase": "Starting …"}
            threading.Thread(target=run_job,
                             args=(job_id, paths[0], paths[1], names[0], names[1]),
                             daemon=True).start()
            return self._json(200, {"job": job_id})
        except Exception as e:
            return self._json(500, {"error": f"{type(e).__name__}: {e}"})


def port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", port)) == 0


def main():
    parser = argparse.ArgumentParser(description="Local DocDiff web application")
    parser.add_argument("port", nargs="?", type=int, default=DEFAULT_PORT)
    parser.add_argument(
        "--no-browser",
        action="store_true",
        help="Start the server without opening the default browser",
    )
    args = parser.parse_args()
    port = args.port
    url = f"http://127.0.0.1:{port}"

    if port_in_use(port):
        print(f"DocDiff already running (or port {port} busy) — opening {url}")
        if not args.no_browser:
            webbrowser.open(url)
        return

    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"\n  DocDiff running at {url}")
    print("  Press Ctrl+C to stop.\n")
    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  Stopped.")


if __name__ == "__main__":
    main()
