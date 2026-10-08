# -*- coding: utf-8 -*-
"""Build a single-file mobile lookup page for Omer second shipment (025/204).

Reads supplier_documents/4/025-204 second shipment_*.xlsx (sheet Sayfa1) via the
same parser as the carton-label script, and writes carton_lookup/index.html with
the carton data embedded as JSON (no server, no prices, works offline).

Usage:
    python scripts/build_carton_lookup_site.py            # build only
    python scripts/build_carton_lookup_site.py --deploy   # build + push to GitHub Pages
"""
from __future__ import annotations

import datetime as dt
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

from build_omer_second_shipment_carton_labels_pdf import (  # noqa: E402
    FLANNEL_MODELS,
    MODEL_NAMES,
    _find_packing_list,
    load_cartons,
)

OUT_DIR = ROOT / "carton_lookup"
OUT_FILE = OUT_DIR / "index.html"

PAGES_REPO = "levynchi/carton-lookup"
PAGES_URL = "https://levynchi.github.io/carton-lookup/"

BRAND_LABELS = {
    "ARYE": "אריה",
    "BABY BASIC": "בייבי בייסיק",
}


def _size_sort_key(size: str):
    digits = "".join(ch for ch in size.split("-")[0] if ch.isdigit())
    return (int(digits) if digits else 0, size)


def build_ranges(cartons: list[dict]) -> list[dict]:
    """Collapse consecutive cartons with same brand/model/size into ranges."""
    ranges: list[dict] = []
    for c in sorted(cartons, key=lambda x: x["carton"]):
        key = (c["brand"], c["model"], c["size"])
        if ranges and ranges[-1]["_key"] == key and ranges[-1]["to"] == c["carton"] - 1:
            ranges[-1]["to"] = c["carton"]
            ranges[-1]["cartons"] += 1
            ranges[-1]["pieces"] += c["pieces"]
        else:
            ranges.append({
                "_key": key,
                "brand": c["brand"],
                "model": c["model"],
                "size": c["size"],
                "from": c["carton"],
                "to": c["carton"],
                "cartons": 1,
                "pieces": c["pieces"],
            })
    for r in ranges:
        r.pop("_key")
    return ranges


def build_html(cartons: list[dict], source_name: str, updated: str) -> str:
    data = {
        "updated": updated,
        "source": source_name,
        "brands": BRAND_LABELS,
        "models": {str(k): v for k, v in MODEL_NAMES.items()},
        "flannel": sorted(FLANNEL_MODELS),
        "cartons": [
            {
                "n": c["carton"],
                "brand": c["brand"],
                "model": c["model"],
                "size": c["size"],
                "packets": c["packets"],
                "pieces": c["pieces"],
            }
            for c in sorted(cartons, key=lambda x: x["carton"])
        ],
        "ranges": build_ranges(cartons),
    }
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    payload = payload.replace("</", "<\\/")
    return TEMPLATE.replace("__DATA__", payload)


TEMPLATE = r"""<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>חיפוש קרטון - משלוח 025/204</title>
<meta name="theme-color" content="#1f2937">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="קרטונים">
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' rx='14' fill='%231f2937'/%3E%3Cpath d='M14 24l18-9 18 9v18l-18 9-18-9z' fill='none' stroke='%23fbbf24' stroke-width='4' stroke-linejoin='round'/%3E%3Cpath d='M14 24l18 9 18-9M32 33v18' fill='none' stroke='%23fbbf24' stroke-width='4'/%3E%3C/svg%3E">
<link rel="apple-touch-icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' fill='%231f2937'/%3E%3Cpath d='M14 24l18-9 18 9v18l-18 9-18-9z' fill='none' stroke='%23fbbf24' stroke-width='4' stroke-linejoin='round'/%3E%3Cpath d='M14 24l18 9 18-9M32 33v18' fill='none' stroke='%23fbbf24' stroke-width='4'/%3E%3C/svg%3E">
<style>
  :root {
    --bg: #f3f4f6; --card: #ffffff; --text: #111827; --muted: #6b7280; --line: #e5e7eb;
    --arye: #1d4ed8; --arye-bg: #dbeafe; --arye-dark: #1e3a8a;
    --baby: #be185d; --baby-bg: #fce7f3; --baby-dark: #831843;
    --warn: #b45309; --warn-bg: #fef3c7;
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; background: var(--bg); color: var(--text);
    font-family: -apple-system, "Segoe UI", Roboto, "Heebo", Arial, sans-serif; }
  body { min-height: 100vh; padding: env(safe-area-inset-top) env(safe-area-inset-right) env(safe-area-inset-bottom) env(safe-area-inset-left); }
  header { background: #1f2937; color: #fff; padding: 8px 14px 7px; }
  header h1 { margin: 0; font-size: 15px; font-weight: 700; }
  header p { margin: 2px 0 0; font-size: 11px; color: #d1d5db; }
  main { max-width: 520px; margin: 0 auto; padding: 8px 12px 16px; }
  .search { position: relative; }
  .search input {
    width: 100%; font-size: 32px; font-weight: 800; text-align: center; letter-spacing: 2px;
    padding: 8px 48px; border: 2px solid var(--line); border-radius: 14px;
    background: var(--card); color: var(--text); outline: none; direction: ltr;
    -webkit-appearance: none; appearance: none;
  }
  .search input:focus { border-color: #9ca3af; }
  .search input::placeholder { color: #d1d5db; font-weight: 600; letter-spacing: 0; font-size: 22px; }
  .search button {
    position: absolute; top: 50%; right: 12px; transform: translateY(-50%);
    width: 40px; height: 40px; border-radius: 50%; border: none; background: var(--line);
    color: var(--muted); font-size: 22px; line-height: 1; cursor: pointer; display: none;
  }
  .search.has-value button { display: block; }
  .hint { text-align: center; color: var(--muted); font-size: 12px; margin: 6px 0 0; }
  body.has-result .hint { display: none; }

  .card { margin-top: 8px; border-radius: 16px; padding: 10px 12px 12px; background: var(--card);
    border: 2px solid var(--line); box-shadow: 0 6px 20px rgba(0,0,0,.06); display: none; }
  .card.show { display: block; animation: pop .18s ease-out; }
  @keyframes pop { from { transform: scale(.97); opacity: .4; } to { transform: scale(1); opacity: 1; } }
  .card.arye { background: var(--arye-bg); border-color: var(--arye); }
  .card.baby { background: var(--baby-bg); border-color: var(--baby); }
  .card.missing { background: var(--warn-bg); border-color: var(--warn); text-align: center; }
  .card.missing .big { font-size: 22px; font-weight: 700; color: var(--warn); }
  .card.missing .sub { color: var(--muted); margin-top: 6px; font-size: 14px; }

  .size-hero { background: rgba(255,255,255,.85); border-radius: 14px; padding: 6px 10px 8px; text-align: center; }
  .size-hero .lbl { font-size: 13px; color: var(--muted); font-weight: 700; }
  .size-hero .val { font-size: 64px; font-weight: 900; line-height: .95; margin-top: 0; direction: ltr; letter-spacing: 1px; }
  .who { display: flex; align-items: baseline; justify-content: space-between; gap: 8px; margin-top: 8px; }
  .brand { font-size: 26px; font-weight: 900; line-height: 1; }
  .arye .brand { color: var(--arye-dark); }
  .baby .brand { color: var(--baby-dark); }
  .carton-no { font-size: 15px; color: var(--muted); font-weight: 700; }
  .carton-no b { color: var(--text); }
  .model-name { margin-top: 4px; font-size: 18px; font-weight: 800; line-height: 1.2; }
  .model-name .tag { display: inline-block; margin-inline-start: 6px; font-size: 12px; font-weight: 700; padding: 2px 8px; border-radius: 999px; background: rgba(255,255,255,.85); border: 1px solid var(--line); color: var(--muted); vertical-align: middle; }
  body.keyboard details, body.keyboard footer { display: none; }

  details { margin-top: 26px; background: var(--card); border-radius: 16px; border: 1px solid var(--line); }
  summary { cursor: pointer; padding: 14px 16px; font-weight: 700; font-size: 15px; list-style: none; display: flex; justify-content: space-between; align-items: center; }
  summary::-webkit-details-marker { display: none; }
  summary::after { content: "▾"; color: var(--muted); transition: transform .15s; }
  details[open] summary::after { transform: rotate(180deg); }
  table { width: 100%; border-collapse: collapse; font-size: 14px; }
  th, td { padding: 9px 8px; text-align: right; border-top: 1px solid var(--line); }
  th { color: var(--muted); font-weight: 600; font-size: 12px; background: #fafafa; }
  td.num { direction: ltr; text-align: right; font-variant-numeric: tabular-nums; }
  tr.arye td:first-child { border-right: 4px solid var(--arye); }
  tr.baby td:first-child { border-right: 4px solid var(--baby); }
  tr.clickable { cursor: pointer; }
  tr.clickable:active { background: #f3f4f6; }
  .brand-head td { background: #f9fafb; font-weight: 800; font-size: 14px; }
  .brand-head.arye td { color: var(--arye-dark); }
  .brand-head.baby td { color: var(--baby-dark); }
  footer { text-align: center; color: var(--muted); font-size: 12px; margin-top: 24px; }
</style>
</head>
<body>
<header>
  <h1>חיפוש קרטון - משלוח עומר 025 / 204</h1>
  <p id="meta"></p>
</header>
<main>
  <div class="search" id="search">
    <input id="q" type="text" inputmode="numeric" pattern="[0-9Xx ]*" placeholder="הקלד מספר קרטון" autofocus autocomplete="off" aria-label="מספר קרטון">
    <button id="clear" type="button" aria-label="נקה">×</button>
  </div>
  <p class="hint" id="hint"></p>

  <section class="card" id="card" aria-live="polite"></section>

  <details id="ranges">
    <summary>טבלת טווחים - איזה קרטונים שייכים למי</summary>
    <table>
      <thead><tr><th>קרטונים</th><th>דגם</th><th>מידה</th><th>כמות</th></tr></thead>
      <tbody id="ranges-body"></tbody>
    </table>
  </details>

  <footer id="footer"></footer>
</main>

<script id="data" type="application/json">__DATA__</script>
<script>
(function () {
  var DATA = JSON.parse(document.getElementById('data').textContent);
  var byNo = {};
  DATA.cartons.forEach(function (c) { byNo[c.n] = c; });
  var nums = DATA.cartons.map(function (c) { return c.n; });
  var minNo = Math.min.apply(null, nums), maxNo = Math.max.apply(null, nums);
  var flannel = {};
  DATA.flannel.forEach(function (m) { flannel[m] = true; });

  function brandClass(b) { return b === 'ARYE' ? 'arye' : 'baby'; }
  function brandLabel(b) { return DATA.brands[b] || b; }
  function modelCode(m) { return String(m).padStart(3, '0'); }
  function modelName(m) { return DATA.models[String(m)] || ''; }
  var input = document.getElementById('q');
  var search = document.getElementById('search');
  var card = document.getElementById('card');
  var hint = document.getElementById('hint');
  var debounceTimer = null;
  var DEBOUNCE_MS = 550;

  document.getElementById('meta').textContent = 'עודכן ' + DATA.updated + ' · ' + DATA.cartons.length + ' קרטונים (' + minNo + 'X–' + maxNo + 'X)';
  document.getElementById('footer').textContent = 'מקור: ' + DATA.source;
  hint.textContent = 'אחרי ההקלדה השדה מתנקה — אפשר להקליד מספר חדש';

  function parseNo(raw) {
    var digits = String(raw || '').replace(/\D/g, '');
    return digits ? parseInt(digits, 10) : null;
  }

  function syncClearBtn() {
    search.classList.toggle('has-value', input.value.trim().length > 0);
  }

  function showCard(n) {
    if (n === null) return;
    var c = byNo[n];
    if (!c) {
      card.className = 'card missing show';
      card.innerHTML = '<div class="big">קרטון ' + n + 'X לא נמצא</div>' +
        '<div class="sub">במשלוח הזה יש קרטונים ' + minNo + 'X עד ' + maxNo + 'X בלבד</div>';
      document.body.classList.add('has-result');
      return;
    }
    var cls = brandClass(c.brand);
    var tags = '<span class="tag">דגם ' + modelCode(c.model) + '</span>';
    tags += flannel[c.model] ? '<span class="tag">פלנל</span>' : '<span class="tag">לא פלנל</span>';
    card.className = 'card ' + cls + ' show';
    card.innerHTML =
      '<div class="size-hero"><div class="lbl">מידה</div><div class="val">' + c.size + '</div></div>' +
      '<div class="who"><div class="brand">' + brandLabel(c.brand) + '</div>' +
        '<div class="carton-no">קרטון <b>' + c.n + 'X</b></div></div>' +
      '<div class="model-name">' + modelName(c.model) + tags + '</div>';
    document.body.classList.add('has-result');
    keepSizeVisible();
  }

  function clearInputKeepCard() {
    input.value = '';
    syncClearBtn();
    input.focus();
  }

  function scheduleClear() {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(clearInputKeepCard, DEBOUNCE_MS);
  }

  function onType() {
    syncClearBtn();
    var n = parseNo(input.value);
    if (n === null) return;
    showCard(n);
    scheduleClear();
  }

  input.addEventListener('input', onType);
  input.addEventListener('keydown', function (e) {
    if (e.key === 'Enter') {
      e.preventDefault();
      var n = parseNo(input.value);
      if (n !== null) showCard(n);
      clearTimeout(debounceTimer);
      clearInputKeepCard();
    }
  });
  document.getElementById('clear').addEventListener('click', function () {
    clearTimeout(debounceTimer);
    clearInputKeepCard();
  });

  var body = document.getElementById('ranges-body');
  var lastBrand = null;
  var html = '';
  DATA.ranges.forEach(function (r) {
    var cls = brandClass(r.brand);
    if (r.brand !== lastBrand) {
      html += '<tr class="brand-head ' + cls + '"><td colspan="4">' + brandLabel(r.brand) + '</td></tr>';
      lastBrand = r.brand;
    }
    var span = r.from === r.to ? (r.from + 'X') : (r.from + 'X–' + r.to + 'X');
    html += '<tr class="clickable ' + cls + '" data-from="' + r.from + '">' +
      '<td class="num">' + span + '</td>' +
      '<td>' + modelCode(r.model) + ' ' + modelName(r.model) + '</td>' +
      '<td class="num">' + r.size + '</td>' +
      '<td class="num">' + r.cartons + '</td></tr>';
  });
  body.innerHTML = html;
  body.addEventListener('click', function (e) {
    var tr = e.target.closest('tr[data-from]');
    if (!tr) return;
    var n = parseInt(tr.getAttribute('data-from'), 10);
    showCard(n);
    clearTimeout(debounceTimer);
    clearInputKeepCard();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  function keepSizeVisible() {
    var hero = card.querySelector('.size-hero') || card;
    if (!card.classList.contains('show')) return;
    var vv = window.visualViewport;
    var visibleBottom = (vv ? vv.offsetTop + vv.height : window.innerHeight) - 8;
    var rect = hero.getBoundingClientRect();
    if (rect.bottom > visibleBottom) window.scrollBy(0, rect.bottom - visibleBottom);
  }

  function onViewport() {
    var vv = window.visualViewport;
    var keyboard = vv && (window.innerHeight - vv.height) > 80;
    document.body.classList.toggle('keyboard', !!keyboard);
    keepSizeVisible();
  }
  if (window.visualViewport) {
    window.visualViewport.addEventListener('resize', onViewport);
    window.visualViewport.addEventListener('scroll', onViewport);
  }

  var m = location.search.match(/[?&]c=(\d+)/);
  if (m) showCard(parseInt(m[1], 10));
})();
</script>
</body>
</html>
"""


def build() -> Path:
    packing = _find_packing_list()
    cartons = load_cartons(packing)
    updated = dt.datetime.fromtimestamp(packing.stat().st_mtime).strftime("%d.%m.%Y")
    html = build_html(cartons, packing.name, updated)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    OUT_FILE.write_text(html, encoding="utf-8")
    print(f"source: {packing.name} cartons={len(cartons)}")
    print(f"wrote: {OUT_FILE} ({OUT_FILE.stat().st_size:,} bytes)")
    return OUT_FILE


def deploy() -> None:
    """Push carton_lookup/index.html to the GitHub Pages repo (main branch)."""
    tmp = Path(tempfile.mkdtemp(prefix="carton_lookup_"))
    try:
        run = lambda *args: subprocess.run(args, cwd=tmp, check=True, text=True, capture_output=True)  # noqa: E731
        subprocess.run(["git", "clone", "--depth", "1", f"https://github.com/{PAGES_REPO}.git", str(tmp)],
                       check=True, text=True, capture_output=True)
        shutil.copy2(OUT_FILE, tmp / "index.html")
        (tmp / ".nojekyll").touch()
        run("git", "add", "-A")
        status = run("git", "status", "--porcelain").stdout.strip()
        if not status:
            print("deploy: nothing changed")
            return
        run("git", "commit", "-m", f"Update carton lookup {dt.date.today():%d.%m.%Y}")
        run("git", "push", "origin", "HEAD:main")
        print(f"deployed: {PAGES_URL}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    build()
    if "--deploy" in sys.argv:
        deploy()
