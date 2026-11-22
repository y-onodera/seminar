# -*- coding: utf-8 -*-
"""deck/project のスライドを、エクスプローラーからダブルクリックで開ける HTML に書き出す。

- 1枚ごとに1つの HTML ファイル（CSS は各ファイルに埋め込み）
- 同じフォルダーに全ページがある前提で、前後のページと目次へハイパーリンクでつなぐ
- スピーカーノートは別のテキストファイルにまとめる

実行：python -I export_html.py
"""
import html
import json
import re
from pathlib import Path

HERE = Path(__file__).parent
SRC = HERE / "deck" / "project"
OUT = HERE / "第1回_スライド"
NOTES = HERE / "第1回_スライドノート.txt"

deck = json.loads((SRC / "deck.json").read_text(encoding="utf-8"))
deck_css = (SRC / "deck.css").read_text(encoding="utf-8")
# Web フォントが読めない環境（オフライン）でも日本語フォントで表示する
deck_css = deck_css.replace("'Noto Sans JP', Arial", "'Noto Sans JP', 'Yu Gothic', Meiryo, Arial")
font_links = "".join(
    f'<link rel="stylesheet" href="{html.escape(f["href"])}">' for f in deck["faces"].values() if "href" in f
)

BASE_CSS = """
* { margin: 0; }
html, body { height: 100%; }
body { font: 32px/1.2 sans-serif; background: #2B3033; display: flex; flex-direction: column; align-items: center; justify-content: center; }
.stage { width: 1920px; height: 1080px; box-shadow: 0 8px 40px rgba(0,0,0,.4); }
.stage > section { width: 1920px; height: 1080px; overflow: hidden; box-sizing: border-box; position: relative; isolation: isolate; }
aside { display: none !important; }
h1 { font-size: 96px; font-weight: 600 } h2 { font-size: 64px; font-weight: 600 } h3 { font-size: 44px; font-weight: 600 }
p, li, h1, h2, h3, td, th { overflow-wrap: break-word; text-wrap: pretty; }
table { width: 100%; border-collapse: separate; border-spacing: 0; border: solid #0003; border-width: 2px 0 0 2px; }
th, td { padding: .35em .6em; text-align: left; border: solid #0003; border-width: 0 2px 2px 0; }
x-shape, x-icon, x-embed { display: block; }
x-shape[kind="arrow-right"] { clip-path: polygon(0 30%, 60% 30%, 60% 0, 100% 50%, 60% 100%, 60% 70%, 0 70%); }
x-shape[kind="arrow-down"] { clip-path: polygon(30% 0, 70% 0, 70% 60%, 100% 60%, 50% 100%, 0 60%, 30% 60%); }
.nav { display: flex; gap: 28px; align-items: center; margin-top: 18px; font: 20px 'Yu Gothic', Meiryo, sans-serif; color: #C9D3D6; }
.nav a { color: #E8EEF0; text-decoration: none; padding: 6px 16px; border: 1px solid #5E6A6F; border-radius: 8px; }
.nav a:hover { background: #3C4447; }
.nav .off { color: #5E6A6F; padding: 6px 16px; border: 1px solid #3C4447; border-radius: 8px; }
"""

# 画面に合わせて 1920×1080 のスライドを縮小する（ナビゲーションは縮小しない）
FIT_JS = """
<script>
function fit() {
  var s = Math.min(window.innerWidth / 1920, (window.innerHeight - 90) / 1080);
  document.querySelector('.stage').style.zoom = s;
}
window.addEventListener('resize', fit); fit();
document.addEventListener('keydown', function (e) {
  var id = { ArrowRight: 'next', PageDown: 'next', ' ': 'next', ArrowLeft: 'prev', PageUp: 'prev' }[e.key];
  var a = id && document.getElementById(id);
  if (a && a.href) { location.href = a.href; }
});
</script>
"""

order = deck["order"]
names = [f"{i + 1:02d}_{sid}.html" for i, sid in enumerate(order)]
OUT.mkdir(exist_ok=True)
for old in OUT.glob("*.html"):
    old.unlink()

titles, notes = [], []
for i, sid in enumerate(order):
    raw = (SRC / "slides" / f"{sid}.html").read_text(encoding="utf-8")
    m = re.search(r"<aside>(.*?)</aside>", raw, re.S)
    notes.append(html.unescape(m.group(1).strip()) if m else "")
    t = re.search(r"<h[12][^>]*>(.*?)</h[12]>", raw, re.S)
    title = html.unescape(re.sub(r"<[^>]+>", "", t.group(1))).strip() if t else sid
    titles.append(title)

    prev_link = f'<a id="prev" href="{names[i - 1]}">◀ 前へ</a>' if i > 0 else '<span class="off">◀ 前へ</span>'
    next_link = f'<a id="next" href="{names[i + 1]}">次へ ▶</a>' if i < len(order) - 1 else '<span class="off">次へ ▶</span>'
    page = f"""<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<title>{i + 1:02d} {html.escape(title)} | {html.escape(deck["title"])}</title>
{font_links}
<style>{BASE_CSS}</style>
<style>body {{ counter-reset: page {i}; }}
{deck_css}</style>
</head><body>
<div class="stage">
{raw}
</div>
<nav class="nav">{prev_link}<span>{i + 1} / {len(order)}</span><a href="00_目次.html">目次</a>{next_link}</nav>
{FIT_JS}
</body></html>
"""
    (OUT / names[i]).write_text(page, encoding="utf-8")

# 目次
section_starts = {v["start"]: v["description"] for v in deck["sections"].values()}
items = []
for i, sid in enumerate(order):
    if sid in section_starts:
        items.append(f'<h2>{html.escape(section_starts[sid])}</h2>')
    items.append(f'<p><a href="{names[i]}">{i + 1:02d}　{html.escape(titles[i])}</a></p>')
index = f"""<!doctype html>
<html lang="ja"><head><meta charset="utf-8"><title>目次 | {html.escape(deck["title"])}</title>
<style>
body {{ font-family: 'Yu Gothic', Meiryo, sans-serif; background: #F7F8F6; color: #1E292E; margin: 0; padding: 48px 64px; }}
h1 {{ font-size: 32px; color: #16343A; margin: 0 0 8px; }}
h2 {{ font-size: 20px; color: #0F6E64; margin: 28px 0 8px; }}
p {{ margin: 4px 0; font-size: 18px; }}
a {{ color: #1E292E; }}
.start {{ display: inline-block; margin-top: 16px; padding: 10px 24px; background: #16343A; color: #FBFBF8; border-radius: 8px; text-decoration: none; }}
</style></head><body>
<h1>{html.escape(deck["title"])}</h1>
<a class="start" href="{names[0]}">はじめから表示する ▶</a>
{''.join(items)}
</body></html>
"""
(OUT / "00_目次.html").write_text(index, encoding="utf-8")

# ノート
lines = [f"{deck['title']}　スピーカーノート", ""]
for i, sid in enumerate(order):
    if sid in section_starts:
        lines += ["=" * 40, f"■ {section_starts[sid]}", "=" * 40, ""]
    lines += [f"【{i + 1:02d}】{titles[i]}（{names[i]}）", notes[i] or "（ノートなし）", ""]
NOTES.write_text("\n".join(lines), encoding="utf-8-sig")

print(f"{len(order)} pages + index -> {OUT}")
print(f"notes -> {NOTES}")
