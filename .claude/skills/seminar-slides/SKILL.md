---
name: seminar-slides
description: 固定資産仕訳セミナー（このリポジトリ）のHTMLスライド（deck/project/）を作る・直す・書き出す・公開するときの技術的な手順。スライドの形式、レイアウトの決まり、はみ出しの計測、export_html.py での書き出し、claude.ai の Slides 型 Artifact の更新方法。文章の書き方と用語は seminar-writing に従う。
---

# 固定資産仕訳セミナーのスライドの作り方

スライドの中身は骨子（`00_骨子.md`）の「スライドに反映する内容」に合わせ、言葉は seminar-writing に従う。ここには技術的な手順だけを書く。スライドとノートの分け方（情報を詰め込まない）は seminar-writing の「スライドへの載せ方」に従う。

## ファイルの形式（claude.ai の Slides 型 Artifact）

- `deck/project/deck.json`：デッキの定義
  - `v`（5）、`lists`（"css"）、`format`（`{"v":1,"shared":["project/deck.css"]}`）、`title`
  - `order`：スライドのIDの並び。IDは `slides/<id>.html` のファイル名
  - `sections`：`{"s1":{"description":"…（時間）","start":"<最初のスライドのID>"}, …}`。目次とノートの区切りに使う
  - `faces`：Noto Sans JP（Google Fonts）
- `deck/project/deck.css`：全スライド共通のスタイル
- `deck/project/slides/<id>.html`：1ファイルに `<section id="<id>">` を1つ。話す内容は `<aside>` に書き、先頭に `【X分（累計Y分）】` と時間を付ける（ノートとして書き出される）

## レイアウトの決まり

- 1枚は 1920×1080。`section` の余白は上・左右 128px、下 160px。フッター（`.foot`）は下から 64px の位置に固定され、ページ番号が自動で付く
- 本文の下端は 920px 以内に収める。文字は 24px 以上
- 表の列幅は、最初の行のセルに `style="width:…%"` で指定する
- 主なクラス（`deck.css`）
  - 見出しまわり：`.eyebrow`（部の名前の札）、`h2`、`.lead`（リード文）。`section` に `l1` を付けると緑、`l2` を付けると橙の配色
  - レイアウト：`.row`（横並び）、`.col`（縦並び）、`.f1`〜`.f6`（幅の比）
  - カード：`.card`、`.t0`（灰）・`.t1`（緑）・`.t2`（橙）、`h3` の色 `.c1`・`.c2`
  - 本文：`.pts`（箇条書き）、`.txt`、`.note`（注記）、強調 `.k`（緑）・`.o`（橙）
  - 表：`.tbl`、見出し行 `.hd`・`.hd1`・`.hd2`、行の色 `.ra`・`.r1`・`.r2`・`.r0`、中央寄せ `.c`
  - 図：矢印 `<x-shape kind="arrow-right" class="arr">`・`<x-shape kind="arrow-down" class="arrd">`、`.flowbox`、2×2 の `.grid2`（`.gh` 見出し、`.gc` セル）
  - 表紙・最後：`section` に `cover`・`closing`。`.kicker`・`.sub`・`.badges`・`.badge`（`.b0`・`.b1`・`.b2`）
- 時間配分：ノートの累計の最後が 60分 になるようにし、`deck.json` の `sections` の時間とそろえる

## 書き出し（ローカル配布用・GitHub Pages）

- セミナーではローカルPCで HTML ファイルをダブルクリックして開く。書き出したHTMLは、サーバーなしで動くようにする（CSS・スクリプトは埋め込み、fetch や外部ファイルの読み込みに頼らない。フォントは読み込めなくても代わりのフォントで表示される）

- `python -I export_html.py` で、`第1回_スライド/`（1枚1ファイルのHTML。前後・目次へのリンク付きで、ダブルクリックで開ける）、`00_目次.html`、`第1回_スライドノート.txt`（`<aside>` を集めたもの）を書き出す
- 書き出したHTMLは CSS を埋め込み、画面の大きさに合わせて縮小表示する。フォントは Noto Sans JP がなければ Yu Gothic・Meiryo を使う
- GitHub Pages は main の `第1回_スライド/` をそのまま表示する。スライドを直したら書き出し直してコミットする

## はみ出しの計測

スクリーンショットは環境によって撮れないことがあるので、数値で確かめる。

1. 書き出したフォルダをローカルのHTTPサーバーで開く（例：`python -m http.server 8765 --bind 127.0.0.1 --directory <リポジトリのフォルダ>`。Claude Code のデスクトップ版では `.claude/launch.json` に登録してプレビューで開く）
2. `第1回_スライド/00_目次.html` を開き、次のスクリプトを実行する。各スライドを 1920×1170 の iframe で読み込み（縮小率が1になる）、本文の下端と、見出しの折り返しを出す

```js
const links = [...new Set([...document.querySelectorAll('a')].map(a => a.getAttribute('href')).filter(h => /^\d\d_.*\.html$/.test(h) && !h.startsWith('00_')))];
const lines = (d, el) => { const r = d.createRange(); r.selectNodeContents(el); return new Set([...r.getClientRects()].filter(x => x.width > 1).map(x => Math.round(x.top))).size; };
const out = [];
for (const f of links) {
  const ifr = document.createElement('iframe');
  ifr.style.cssText = 'position:absolute;left:-6000px;top:0;width:1920px;height:1170px;border:0';
  ifr.src = f + '?v=' + Date.now();
  document.body.appendChild(ifr);
  await new Promise(r => ifr.onload = r);
  await ifr.contentDocument.fonts.ready;
  await new Promise(r => setTimeout(r, 150));
  const d = ifr.contentDocument, sec = d.querySelector('section'), sr = sec.getBoundingClientRect();
  let maxB = 0;
  for (const el of sec.children) { if (el.matches('.foot, aside')) continue; maxB = Math.max(maxB, el.getBoundingClientRect().bottom - sr.top); }
  const wraps = [...sec.querySelectorAll('th, h2, .eyebrow, h1, .badge')].filter(t => lines(d, t) > 1).map(t => t.textContent.trim().slice(0, 15));
  out.push(f + ' ' + Math.round(maxB) + (maxB > 920 ? ' OVER' : '') + (wraps.length ? ' wrap:' + wraps.join('|') : ''));
  ifr.remove();
}
out.join('\n');
```

- 920 を超えたスライド、表の見出しや `h2` が折り返したスライドを直す。`.grid2` を使うスライドは、余白いっぱいに広がるので 920 ちょうどになる
- 直し方の例：文を短くする、表の列幅を配分し直す、カードの箇条書きを減らす

## claude.ai の Artifact の更新

- Slides 型の Artifact（非公開）：https://claude.ai/artifact/TDU8MVFwJpCbZsY96Pn6KT
- Artifact ツールの publish で、`url` にこのリンク、`file_path` に `deck/project/deck.json`、`root` に `deck` フォルダ、`files` に変えたスライド（`project/slides/<id>.html`）を渡す。スライドを消すときは `files` でそのパスに `null` を渡す
- 前回の公開のあとに別の版が出ていると、公開が止められる。そのときは Artifact を読み直し、差分が自分の変更だけかを確かめてから公開し直す
- 公開はローカルの Claude Code から行っている。claude.ai のプロジェクトから行う場合は、そこで Artifact ツールが使えるかを先に確かめる

## このPCの環境の注意

- 既定の Python は 3.15 のアルファ版で、lxml などが動かない。`export_html.py` は標準ライブラリだけなので動く。旧PowerPoint版の `build_deck.py`（python-pptx を使う）を動かすときは、uv で Python 3.13 を使った
