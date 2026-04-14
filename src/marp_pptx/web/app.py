"""Flask web UI for marp-pptx."""
from __future__ import annotations

import hashlib
import shutil
import subprocess
import tempfile
import uuid
from pathlib import Path

from flask import Flask, request, send_file, render_template_string, jsonify, redirect, url_for


# Cache for rendered slide thumbnails (keyed by MD content hash + settings)
_PREVIEW_CACHE_DIR = Path(tempfile.gettempdir()) / "marp_pptx_previews"
_PREVIEW_CACHE_DIR.mkdir(exist_ok=True)

_SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
_PDFTOPPM = shutil.which("pdftoppm")


def _render_pptx_to_pngs(pptx_path: Path, out_dir: Path, dpi: int = 100) -> list[Path]:
    """Convert a PPTX file to per-slide PNG thumbnails using soffice + pdftoppm.

    Returns the sorted list of PNG paths. Returns [] if tools unavailable
    or conversion fails.
    """
    if _SOFFICE is None or _PDFTOPPM is None:
        return []
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        # Step 1: PPTX → PDF via LibreOffice
        subprocess.run(
            [_SOFFICE, "--headless", "--convert-to", "pdf",
             "--outdir", str(out_dir), str(pptx_path)],
            check=True, capture_output=True, timeout=60,
        )
        pdf = out_dir / (pptx_path.stem + ".pdf")
        if not pdf.exists():
            return []
        # Step 2: PDF → PNG per page via pdftoppm
        subprocess.run(
            [_PDFTOPPM, "-png", "-r", str(dpi), str(pdf), str(out_dir / "slide")],
            check=True, capture_output=True, timeout=60,
        )
        return sorted(out_dir.glob("slide-*.png"))
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        return []


EDITOR_HTML = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>marp-pptx Editor</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
html, body { height: 100%; }
body { font-family: -apple-system, 'Segoe UI', 'Hiragino Sans', sans-serif; background: #f7f7f7; color: #1a1a1a; }
.topbar { background: #1a1a1a; color: white; padding: 10px 20px; display: flex; align-items: center; gap: 16px; }
.topbar h1 { font-size: 1.1em; font-weight: 600; }
.topbar a { color: #bbb; text-decoration: none; font-size: 0.9em; }
.topbar a:hover { color: white; }
.topbar .spacer { flex: 1; }
.layout { display: grid; grid-template-columns: 1fr 1fr 300px; height: calc(100vh - 42px); }
.editor-pane { display: flex; flex-direction: column; border-right: 1px solid #ddd; }
.preview-pane { background: #eee; overflow-y: auto; padding: 16px; border-right: 1px solid #ddd; }
.preview-pane h3 { font-size: 0.85em; text-transform: uppercase; letter-spacing: 0.05em; color: #666; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center; }
.preview-pane .slide-thumb { background: white; margin-bottom: 12px; border: 1px solid #ddd; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
.preview-pane .slide-thumb img { width: 100%; display: block; border-radius: 4px 4px 0 0; }
.preview-pane .slide-thumb .caption { padding: 6px 10px; font-size: 0.75em; color: #999; border-top: 1px solid #eee; }
.preview-empty { color: #999; font-size: 0.85em; text-align: center; padding: 40px 20px; }
.preview-loading { color: #666; font-size: 0.85em; text-align: center; padding: 40px 20px; }
.preview-btn { background: #f0f0f0; border: 1px solid #ccc; padding: 4px 10px; border-radius: 3px; font-size: 0.75em; cursor: pointer; }
.preview-btn:hover { background: #e0e0e0; }
.editor-toolbar { background: white; padding: 10px 14px; border-bottom: 1px solid #eee; display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
.editor-toolbar button { background: #f0f0f0; border: 1px solid #ddd; padding: 6px 12px; border-radius: 3px; font-size: 0.85em; cursor: pointer; }
.editor-toolbar button:hover { background: #e0e0e0; }
.editor-toolbar button.primary { background: #1a1a1a; color: white; border-color: #1a1a1a; font-weight: 600; }
.editor-toolbar button.primary:hover { background: #333; }
.editor-toolbar .sep { color: #ccc; margin: 0 4px; }

/* Modal */
.modal-bg { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); display: none; align-items: center; justify-content: center; z-index: 100; }
.modal-bg.open { display: flex; }
.modal { background: white; border-radius: 8px; max-width: 720px; width: 90%; max-height: 85vh; overflow: hidden; display: flex; flex-direction: column; }
.modal-header { padding: 16px 24px; border-bottom: 1px solid #eee; display: flex; align-items: center; justify-content: space-between; }
.modal-header h2 { font-size: 1.1em; }
.modal-header .close { background: none; border: none; font-size: 1.3em; cursor: pointer; color: #999; }
.modal-body { padding: 20px 24px; overflow-y: auto; flex: 1; }
.modal-footer { padding: 14px 24px; border-top: 1px solid #eee; display: flex; gap: 8px; justify-content: flex-end; background: #f9f9f9; }
.modal-footer button { padding: 8px 18px; border-radius: 3px; border: 1px solid #ddd; background: white; cursor: pointer; font-size: 0.9em; }
.modal-footer button.primary { background: #1a1a1a; color: white; border-color: #1a1a1a; font-weight: 600; }

/* Type picker */
.type-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 8px; }
.type-card { border: 1px solid #ddd; border-radius: 4px; padding: 10px 12px; cursor: pointer; background: white; }
.type-card:hover { background: #eef; border-color: #88c; }
.type-card .type-name { font-family: ui-monospace, 'SF Mono', monospace; font-size: 0.85em; font-weight: 600; color: #1a1a1a; }
.type-card .type-geom { font-size: 0.9em; margin-top: 4px; color: #444; }
.type-card .type-meaning { font-size: 0.78em; color: #777; margin-top: 2px; }
.type-category-header { grid-column: 1/-1; font-size: 0.75em; font-weight: 600; color: #555; text-transform: uppercase; letter-spacing: 0.05em; margin-top: 10px; padding: 4px 0; border-bottom: 1px solid #eee; }

/* Form fields */
.form-row { margin-bottom: 14px; }
.form-row label { display: block; font-weight: 600; font-size: 0.85em; margin-bottom: 6px; color: #333; }
.form-row input[type="text"], .form-row textarea { width: 100%; padding: 7px 10px; border: 1px solid #ccc; border-radius: 3px; font-size: 0.9em; font-family: inherit; }
.form-row textarea { resize: vertical; min-height: 60px; font-family: ui-monospace, 'SF Mono', monospace; font-size: 0.85em; }
.form-row .hint { font-size: 0.75em; color: #999; margin-top: 3px; }
.form-row input[type="checkbox"] { margin-right: 6px; }
.array-items { display: flex; flex-direction: column; gap: 8px; }
.array-item { display: flex; gap: 8px; align-items: flex-start; padding: 8px; background: #f7f7f7; border-radius: 3px; }
.array-item > div { flex: 1; }
.array-item .remove-btn { background: #fff; border: 1px solid #ccc; color: #c62828; padding: 4px 10px; border-radius: 3px; cursor: pointer; font-size: 0.85em; flex-shrink: 0; }
.array-item .remove-btn:hover { background: #fce4e4; }
.add-item-btn { margin-top: 8px; background: #eef; color: #1a1a1a; border: 1px dashed #88c; padding: 8px 14px; border-radius: 3px; cursor: pointer; font-size: 0.85em; }
.add-item-btn:hover { background: #ccd; }
textarea#md-editor {
    flex: 1; width: 100%; border: none; padding: 16px 20px;
    font-family: ui-monospace, 'SF Mono', Menlo, monospace;
    font-size: 13px; line-height: 1.6; resize: none; outline: none;
    background: white;
}
aside { background: white; padding: 20px; overflow-y: auto; border-left: 1px solid #ddd; }
aside h2 { font-size: 0.95em; margin: 16px 0 8px; color: #555; text-transform: uppercase; letter-spacing: 0.05em; }
aside h2:first-child { margin-top: 0; }
label { display: block; font-weight: 600; margin: 10px 0 4px; font-size: 0.85em; }
select, input[type="text"] { width: 100%; padding: 6px 8px; border: 1px solid #ddd; border-radius: 3px; font-size: 0.9em; }
input[type="range"] { width: 100%; padding: 0; }
.slider-row { display: flex; align-items: center; gap: 8px; }
.slider-val { min-width: 36px; font-variant-numeric: tabular-nums; font-size: 0.85em; color: #666; }
button.primary { width: 100%; background: #1a1a1a; color: white; border: none; padding: 12px; border-radius: 4px; font-size: 0.95em; cursor: pointer; margin-top: 16px; font-weight: 600; }
button.primary:hover { background: #333; }
button.primary:disabled { background: #999; cursor: wait; }
.stats { font-size: 0.8em; color: #666; margin-top: 8px; padding: 10px; background: #f9f9f9; border-radius: 4px; }
.stats span { font-weight: 600; color: #1a1a1a; }
.sample-btn { display: block; width: 100%; text-align: left; padding: 8px 10px; margin-bottom: 4px; background: #f7f7f7; border: 1px solid #eee; border-radius: 3px; font-size: 0.85em; cursor: pointer; color: #555; }
.sample-btn:hover { background: #eef; color: #1a1a1a; }
.status { margin-top: 10px; padding: 8px 10px; border-radius: 3px; font-size: 0.85em; display: none; }
.status.ok { background: #e6f7e6; color: #2e7d32; display: block; }
.status.err { background: #fce4e4; color: #c62828; display: block; }
</style>
</head>
<body>
<div class="topbar">
<h1>marp-pptx Editor</h1>
<a href="/">変換画面に戻る</a>
<a href="/types-page">型一覧</a>
<div class="spacer"></div>
<span style="font-size:0.8em; color:#999">.md 保存不要・ブラウザ内で編集 → PPTX 生成</span>
</div>

<div class="layout">
<div class="editor-pane">
<div class="editor-toolbar">
<button class="primary" onclick="openTypePicker()">+ スライドを追加（型から選ぶ）</button>
<span class="sep">|</span>
<button onclick="insertSnippet('plain')">プレーン</button>
<button onclick="insertSnippet('bullets')">箇条書き</button>
<button onclick="insertSnippet('divider')">区切り</button>
<span class="sep">|</span>
<button onclick="if(confirm('エディタ内容を全削除しますか？')) document.getElementById('md-editor').value=''; updateStats();">全削除</button>
</div>

<!-- Type picker modal -->
<div class="modal-bg" id="picker-modal">
<div class="modal">
<div class="modal-header">
<h2>スライド型を選ぶ</h2>
<button class="close" onclick="closeModal('picker-modal')">×</button>
</div>
<div class="modal-body">
<div class="type-grid" id="type-grid"></div>
</div>
</div>
</div>

<!-- Form modal -->
<div class="modal-bg" id="form-modal">
<div class="modal">
<div class="modal-header">
<h2 id="form-title">型の入力</h2>
<button class="close" onclick="closeModal('form-modal')">×</button>
</div>
<div class="modal-body" id="form-body"></div>
<div class="modal-footer">
<button onclick="closeModal('form-modal')">キャンセル</button>
<button class="primary" onclick="submitForm()">スライドを追加</button>
</div>
</div>
</div>
<textarea id="md-editor" spellcheck="false" placeholder="ここにMarkdownを書くか、右のサンプルからロードしてください"></textarea>
</div>

<div class="preview-pane">
<h3>
<span>プレビュー (実レンダリング)</span>
<button class="preview-btn" onclick="refreshPreview()">更新</button>
</h3>
<div id="preview-content">
<div class="preview-empty">エディタに内容を入れて<br>「更新」を押してください</div>
</div>
</div>

<aside>
<h2>サンプル</h2>
<button class="sample-btn" onclick="loadSample('minimal')">📄 最小雛形</button>
<button class="sample-btn" onclick="loadSample('all')">📚 全型カタログ</button>
<button class="sample-btn" onclick="loadSample('academic')">🎓 学術発表サンプル</button>

<h2>出力設定</h2>
<label>Palette</label>
<select id="palette">
<option value="">Default (mono)</option>
{% for p in palettes %}<option value="{{ p }}">{{ p }}</option>{% endfor %}
</select>

<label>Font Scale</label>
<div class="slider-row">
<input type="range" id="font-scale" min="0.7" max="1.3" step="0.05" value="1.0">
<span class="slider-val" id="fs-val">1.00</span>
</div>

<label>ファイル名</label>
<input type="text" id="output-name" value="slides_editable.pptx">

<button class="primary" id="gen-btn" onclick="generate()">→ PPTX を生成してダウンロード</button>

<div class="status" id="status"></div>
<div class="stats" id="stats"></div>
</aside>
</div>

<script>
const editor = document.getElementById('md-editor');
const stats = document.getElementById('stats');
const fsRange = document.getElementById('font-scale');
const fsVal = document.getElementById('fs-val');
const statusEl = document.getElementById('status');

fsRange.addEventListener('input', () => fsVal.textContent = parseFloat(fsRange.value).toFixed(2));

// Live stats
function updateStats() {
    const t = editor.value;
    if (!t.trim()) { stats.innerHTML = '未入力'; return; }
    const slides = t.split(/\\n---\\n/).filter(x => x.trim()).length;
    const chars = t.length;
    const types = [...t.matchAll(/<!--\\s+_class:\\s+(\\S+)\\s+-->/g)].map(m => m[1]);
    const typeCounts = {};
    types.forEach(x => typeCounts[x] = (typeCounts[x] || 0) + 1);
    const typeList = Object.entries(typeCounts).map(([k,v]) => v > 1 ? `${k}×${v}` : k).join(', ');
    stats.innerHTML = `<span>${slides}</span> slides · <span>${chars}</span> chars${typeList ? '<br>型: ' + typeList : ''}`;
}
editor.addEventListener('input', updateStats);

// ── Type schemas: fields + MD template per type ──
const TYPE_SCHEMAS = {};
let TYPES_META = [];  // loaded from /api/types

async function loadTypeMeta() {
    try {
        const r = await fetch('/api/types');
        TYPES_META = await r.json();
    } catch(e) { console.error('型一覧の取得に失敗', e); }
}

// Helpers for building MD
function esc(s) { return String(s || ''); }
function joinLines(arr) { return arr.filter(Boolean).join('\\n'); }

// Schema: { fields: [...], toMd: (data) => string }
// field: { name, label, type: 'text'|'textarea'|'array'|'checkbox', default, hint, subfields? }

TYPE_SCHEMAS['plain'] = {
    label: 'プレーン本文',
    fields: [
        { name: 'h1', label: 'タイトル (H1)', type: 'text', default: 'タイトル' },
        { name: 'body', label: '本文 (Markdown)', type: 'textarea', default: '- ポイント1\\n- ポイント2' },
    ],
    toMd: d => `# ${esc(d.h1)}\\n${esc(d.body)}`,
};

TYPE_SCHEMAS['title'] = {
    label: 'title — 表紙',
    fields: [
        { name: 'h1', label: 'メインタイトル', type: 'text', default: 'タイトル' },
        { name: 'h2', label: 'サブタイトル', type: 'text', default: 'サブタイトル' },
        { name: 'author', label: '発表者名', type: 'text', default: '' },
        { name: 'date', label: '日付', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: title -->\\n# ${esc(d.h1)}${d.h2 ? '\\n## '+esc(d.h2) : ''}${d.author ? '\\n'+esc(d.author) : ''}${d.date ? '\\n'+esc(d.date) : ''}`,
};

TYPE_SCHEMAS['divider'] = {
    label: 'divider — 章区切り',
    fields: [
        { name: 'h1', label: '章タイトル', type: 'text', default: '第○章' },
        { name: 'h2', label: '章サブ', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: divider -->\\n# ${esc(d.h1)}${d.h2 ? '\\n## '+esc(d.h2) : ''}`,
};

TYPE_SCHEMAS['end'] = {
    label: 'end — 終了',
    fields: [
        { name: 'h1', label: 'メッセージ', type: 'text', default: 'Thank You' },
        { name: 'sub', label: '補足（任意）', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: end -->\\n# ${esc(d.h1)}${d.sub ? '\\n'+esc(d.sub) : ''}`,
};

TYPE_SCHEMAS['cols-2'] = {
    label: 'cols-2 — 2カラム',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '比較' },
        { name: 'left_title', label: '左カラム見出し (H3)', type: 'text', default: '従来' },
        { name: 'left_body', label: '左カラム本文', type: 'textarea', default: '- 項目A\\n- 項目B' },
        { name: 'right_title', label: '右カラム見出し (H3)', type: 'text', default: '提案' },
        { name: 'right_body', label: '右カラム本文', type: 'textarea', default: '- 項目A\\n- 項目B' },
    ],
    toMd: d => `<!-- _class: cols-2 -->\\n# ${esc(d.h1)}\\n<div>\\n\\n### ${esc(d.left_title)}\\n${esc(d.left_body)}\\n</div>\\n<div>\\n\\n### ${esc(d.right_title)}\\n${esc(d.right_body)}\\n</div>`,
};

TYPE_SCHEMAS['cols-3'] = {
    label: 'cols-3 — 3カラム',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '3観点' },
        { name: 'c1_title', label: '列1 見出し', type: 'text', default: '観点1' },
        { name: 'c1_body', label: '列1 本文', type: 'textarea', default: '- A\\n- B' },
        { name: 'c2_title', label: '列2 見出し', type: 'text', default: '観点2' },
        { name: 'c2_body', label: '列2 本文', type: 'textarea', default: '- A\\n- B' },
        { name: 'c3_title', label: '列3 見出し', type: 'text', default: '観点3' },
        { name: 'c3_body', label: '列3 本文', type: 'textarea', default: '- A\\n- B' },
    ],
    toMd: d => `<!-- _class: cols-3 -->\\n# ${esc(d.h1)}\\n<div>\\n\\n### ${esc(d.c1_title)}\\n${esc(d.c1_body)}\\n</div>\\n<div>\\n\\n### ${esc(d.c2_title)}\\n${esc(d.c2_body)}\\n</div>\\n<div>\\n\\n### ${esc(d.c3_title)}\\n${esc(d.c3_body)}\\n</div>`,
};

TYPE_SCHEMAS['sandwich'] = {
    label: 'sandwich — 上・中・下',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'タイトル' },
        { name: 'lead', label: '上部リード文', type: 'textarea', default: '概要を1行で。' },
        { name: 'left_title', label: '中央左 見出し', type: 'text', default: '従来' },
        { name: 'left_body', label: '中央左 本文', type: 'textarea', default: '- 項目A' },
        { name: 'right_title', label: '中央右 見出し', type: 'text', default: '提案' },
        { name: 'right_body', label: '中央右 本文', type: 'textarea', default: '- 項目A' },
        { name: 'conclusion', label: '下部結論', type: 'textarea', default: '**結論:** ...' },
    ],
    toMd: d => `<!-- _class: sandwich -->\\n# ${esc(d.h1)}\\n<div class="top">\\n<div class="lead">${esc(d.lead)}</div>\\n</div>\\n<div class="columns">\\n<div>\\n\\n### ${esc(d.left_title)}\\n${esc(d.left_body)}\\n</div>\\n<div>\\n\\n### ${esc(d.right_title)}\\n${esc(d.right_body)}\\n</div>\\n</div>\\n<div class="bottom">\\n<div class="conclusion">${esc(d.conclusion)}</div>\\n</div>`,
};

TYPE_SCHEMAS['equation'] = {
    label: 'equation — 数式',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '数式' },
        { name: 'formula', label: 'LaTeX式（$$なしで記入）', type: 'textarea', default: 'E = mc^2', hint: '例: \\\\frac{a}{b}, \\\\sum_{i=1}^n x_i' },
        { name: 'vars', label: '変数説明', type: 'array',
          subfields: [
            { name: 'sym', label: '記号 (LaTeX)', type: 'text', default: 'E' },
            { name: 'desc', label: '説明', type: 'text', default: 'エネルギー' },
          ], default: [{sym:'E',desc:'エネルギー'},{sym:'m',desc:'質量'},{sym:'c',desc:'光速'}] },
    ],
    toMd: d => {
        const vars_ = (d.vars||[]).map(v => `<span>$${esc(v.sym)}$</span><span>${esc(v.desc)}</span>`).join('\\n');
        return `<!-- _class: equation -->\\n# ${esc(d.h1)}\\n<div class="eq-main">\\n$$${esc(d.formula)}$$\\n</div>${vars_?`\\n<div class="eq-desc">\\n${vars_}\\n</div>`:''}`;
    },
};

TYPE_SCHEMAS['kpi'] = {
    label: 'kpi — 主要指標',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '主要指標' },
        { name: 'items', label: '指標項目', type: 'array',
          subfields: [
            { name: 'value', label: '値', type: 'text', default: '98%' },
            { name: 'label', label: 'ラベル', type: 'text', default: '精度' },
          ], default: [{value:'98%',label:'精度'},{value:'1.2s',label:'応答'},{value:'10x',label:'高速化'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div><span class="kpi-value">${esc(it.value)}</span><span class="kpi-label">${esc(it.label)}</span></div>`).join('\\n');
        return `<!-- _class: kpi -->\\n# ${esc(d.h1)}\\n<div class="kpi-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['funnel'] = {
    label: 'funnel — 絞り込み',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '絞り込み' },
        { name: 'items', label: 'ステージ', type: 'array',
          subfields: [
            { name: 'label', label: 'ラベル', type: 'text', default: 'ステージ' },
            { name: 'value', label: '値', type: 'text', default: '' },
          ], default: [{label:'応募',value:'1,000'},{label:'書類通過',value:'200'},{label:'面接通過',value:'50'},{label:'採用',value:'10'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div><span class="fn-label">${esc(it.label)}</span><span class="fn-value">${esc(it.value)}</span></div>`).join('\\n');
        return `<!-- _class: funnel -->\\n# ${esc(d.h1)}\\n<div class="fn-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['pros-cons'] = {
    label: 'pros-cons — 賛否',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '賛否' },
        { name: 'pros', label: 'Pros 項目', type: 'array',
          subfields: [{ name: 'text', label: '項目', type: 'text', default: '' }],
          default: [{text:'高速'},{text:'省メモリ'}] },
        { name: 'cons', label: 'Cons 項目', type: 'array',
          subfields: [{ name: 'text', label: '項目', type: 'text', default: '' }],
          default: [{text:'実装コスト高'},{text:'依存多い'}] },
    ],
    toMd: d => {
        const pros = (d.pros||[]).map(it => `<li>${esc(it.text)}</li>`).join('');
        const cons = (d.cons||[]).map(it => `<li>${esc(it.text)}</li>`).join('');
        return `<!-- _class: pros-cons -->\\n# ${esc(d.h1)}\\n<div class="pc-pros">\\n<ul>${pros}</ul>\\n</div>\\n<div class="pc-cons">\\n<ul>${cons}</ul>\\n</div>`;
    },
};

TYPE_SCHEMAS['timeline-h'] = {
    label: 'timeline-h — 横タイムライン',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'タイムライン' },
        { name: 'items', label: '時点', type: 'array',
          subfields: [
            { name: 'year', label: '年/時点', type: 'text', default: '2024' },
            { name: 'text', label: '出来事', type: 'text', default: '' },
            { name: 'highlight', label: '強調', type: 'checkbox', default: false },
          ], default: [{year:'2024',text:'企画'},{year:'2025',text:'開発',highlight:true},{year:'2026',text:'リリース'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div class="tl-h-item${it.highlight?' highlight':''}"><div><span class="tl-h-year">${esc(it.year)}</span><span class="tl-h-text">${esc(it.text)}</span></div></div>`).join('\\n');
        return `<!-- _class: timeline-h -->\\n# ${esc(d.h1)}\\n<div class="tl-h-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['zone-matrix'] = {
    label: 'zone-matrix — 2×2 マトリクス',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '2軸評価' },
        { name: 'x_label', label: 'X軸ラベル', type: 'text', default: '重要度' },
        { name: 'y_label', label: 'Y軸ラベル', type: 'text', default: '緊急度' },
        { name: 'tl_label', label: '左上 ラベル', type: 'text', default: 'A' },
        { name: 'tl_body', label: '左上 本文', type: 'text', default: '' },
        { name: 'tr_label', label: '右上 ラベル', type: 'text', default: 'B' },
        { name: 'tr_body', label: '右上 本文', type: 'text', default: '' },
        { name: 'bl_label', label: '左下 ラベル', type: 'text', default: 'C' },
        { name: 'bl_body', label: '左下 本文', type: 'text', default: '' },
        { name: 'br_label', label: '右下 ラベル', type: 'text', default: 'D' },
        { name: 'br_body', label: '右下 本文', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: zone-matrix -->\\n# ${esc(d.h1)}\\n<div class="zm-xlabel">${esc(d.x_label)}</div>\\n<div class="zm-ylabel">${esc(d.y_label)}</div>\\n<div class="zm-tl"><span class="zm-label">${esc(d.tl_label)}</span><span class="zm-body">${esc(d.tl_body)}</span></div>\\n<div class="zm-tr"><span class="zm-label">${esc(d.tr_label)}</span><span class="zm-body">${esc(d.tr_body)}</span></div>\\n<div class="zm-bl"><span class="zm-label">${esc(d.bl_label)}</span><span class="zm-body">${esc(d.bl_body)}</span></div>\\n<div class="zm-br"><span class="zm-label">${esc(d.br_label)}</span><span class="zm-body">${esc(d.br_body)}</span></div>`,
};

TYPE_SCHEMAS['agenda'] = {
    label: 'agenda — 目次',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '本日の内容' },
        { name: 'items', label: '項目', type: 'array',
          subfields: [{ name: 'text', label: '項目', type: 'text', default: '' }],
          default: [{text:'背景'},{text:'手法'},{text:'結果'},{text:'まとめ'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map((it,i) => `${i+1}. ${esc(it.text)}`).join('\\n');
        return `<!-- _class: agenda -->\\n# ${esc(d.h1)}\\n<div class="agenda-list">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['summary'] = {
    label: 'summary — まとめ',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'まとめ' },
        { name: 'items', label: 'ポイント', type: 'array',
          subfields: [{ name: 'text', label: 'ポイント', type: 'text', default: '' }],
          default: [{text:'提案手法は従来比10倍高速'},{text:'精度は同等'},{text:'OSSとして公開'}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<li>${esc(it.text)}</li>`).join('\\n');
        return `<!-- _class: summary -->\\n# ${esc(d.h1)}\\n<ol class="summary-points">\\n${items}\\n</ol>`;
    },
};

TYPE_SCHEMAS['takeaway'] = {
    label: 'takeaway — キーメッセージ',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'Takeaway' },
        { name: 'main', label: '中央の一文', type: 'textarea', default: '型を選ぶだけで伝わるプレゼンに' },
        { name: 'points', label: '補足ポイント', type: 'array',
          subfields: [{ name: 'text', label: '項目', type: 'text', default: '' }],
          default: [{text:'49種の意味的な型'},{text:'完全編集可能'}] },
    ],
    toMd: d => {
        const pts = (d.points||[]).map(it => `<li>${esc(it.text)}</li>`).join('\\n');
        return `<!-- _class: takeaway -->\\n# ${esc(d.h1)}\\n<div class="ta-main">${esc(d.main)}</div>${pts?`\\n<div class="ta-points">\\n<ul>\\n${pts}\\n</ul>\\n</div>`:''}`;
    },
};

TYPE_SCHEMAS['quote'] = {
    label: 'quote — 引用',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '引用' },
        { name: 'text', label: '引用文', type: 'textarea', default: '' },
        { name: 'source', label: '出典', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: quote -->\\n# ${esc(d.h1)}\\n<div class="qt-text">${esc(d.text)}</div>${d.source?`\\n<div class="qt-source">${esc(d.source)}</div>`:''}`,
};

TYPE_SCHEMAS['definition'] = {
    label: 'definition — 用語定義',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '定義' },
        { name: 'term', label: '用語', type: 'text', default: '' },
        { name: 'body', label: '定義文', type: 'textarea', default: '' },
        { name: 'note', label: '補足', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: definition -->\\n# ${esc(d.h1)}\\n<div class="df-term">${esc(d.term)}</div>\\n<div class="df-body">${esc(d.body)}</div>${d.note?`\\n<div class="df-note">${esc(d.note)}</div>`:''}`,
};

TYPE_SCHEMAS['checklist'] = {
    label: 'checklist — チェックリスト',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'チェックリスト' },
        { name: 'items', label: '項目', type: 'array',
          subfields: [
            { name: 'text', label: '項目', type: 'text', default: '' },
            { name: 'done', label: '完了', type: 'checkbox', default: false },
          ], default: [{text:'要件定義',done:true},{text:'実装',done:false}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<li${it.done?' class="done"':''}>${esc(it.text)}</li>`).join('\\n');
        return `<!-- _class: checklist -->\\n# ${esc(d.h1)}\\n<div class="cl-container">\\n<ul>\\n${items}\\n</ul>\\n</div>`;
    },
};

TYPE_SCHEMAS['highlight'] = {
    label: 'highlight — 強調メッセージ',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '強調' },
        { name: 'text', label: '強調する文', type: 'textarea', default: '' },
    ],
    toMd: d => `<!-- _class: highlight -->\\n# ${esc(d.h1)}\\n<div class="hl-text">${esc(d.text)}</div>`,
};

TYPE_SCHEMAS['rq'] = {
    label: 'rq — 研究課題',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '研究課題' },
        { name: 'main', label: 'メインの問い', type: 'textarea', default: '' },
        { name: 'sub', label: '補足', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: rq -->\\n# ${esc(d.h1)}\\n<div class="rq-main">${esc(d.main)}</div>${d.sub?`\\n<div class="rq-sub">${esc(d.sub)}</div>`:''}`,
};

TYPE_SCHEMAS['code'] = {
    label: 'code — コード',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'コード例' },
        { name: 'lang', label: '言語', type: 'text', default: 'python' },
        { name: 'code', label: 'コード', type: 'textarea', default: 'def hello():\\n    print("hi")' },
        { name: 'desc', label: '説明', type: 'text', default: '' },
    ],
    toMd: d => `<!-- _class: code -->\\n# ${esc(d.h1)}\\n<div class="cd-code">\\n\\n\\\`\\\`\\\`${esc(d.lang)}\\n${esc(d.code)}\\n\\\`\\\`\\\`\\n</div>${d.desc?`\\n<div class="cd-desc">${esc(d.desc)}</div>`:''}`,
};

TYPE_SCHEMAS['zone-flow'] = {
    label: 'zone-flow — フロー',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: 'フロー' },
        { name: 'items', label: 'ステップ', type: 'array',
          subfields: [
            { name: 'label', label: 'ラベル', type: 'text', default: '' },
            { name: 'body', label: '本文', type: 'text', default: '' },
          ], default: [{label:'入力',body:''},{label:'処理',body:''},{label:'出力',body:''}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div><span class="zf-label">${esc(it.label)}</span><span class="zf-body">${esc(it.body)}</span></div>`).join('\\n');
        return `<!-- _class: zone-flow -->\\n# ${esc(d.h1)}\\n<div class="zf-container">\\n${items}\\n</div>`;
    },
};

TYPE_SCHEMAS['steps'] = {
    label: 'steps — 手順',
    fields: [
        { name: 'h1', label: 'タイトル', type: 'text', default: '手順' },
        { name: 'items', label: 'ステップ', type: 'array',
          subfields: [
            { name: 'num', label: 'ステップ番号', type: 'text', default: '1' },
            { name: 'title', label: 'タイトル', type: 'text', default: '' },
            { name: 'body', label: '説明', type: 'text', default: '' },
          ], default: [{num:'1',title:'準備',body:''},{num:'2',title:'実行',body:''},{num:'3',title:'確認',body:''}] },
    ],
    toMd: d => {
        const items = (d.items||[]).map(it => `<div><span class="st-num">${esc(it.num)}</span><span class="st-title">${esc(it.title)}</span><span class="st-body">${esc(it.body)}</span></div>`).join('\\n');
        return `<!-- _class: steps -->\\n# ${esc(d.h1)}\\n<div class="st-container">\\n${items}\\n</div>`;
    },
};

// ── Modal management ──
function openModal(id) { document.getElementById(id).classList.add('open'); }
function closeModal(id) { document.getElementById(id).classList.remove('open'); }

async function openTypePicker() {
    if (TYPES_META.length === 0) await loadTypeMeta();
    const grid = document.getElementById('type-grid');
    const byCategory = {};
    const CAT_ORDER = ['meta','structure','temporal','convergence','evaluation','knowledge','flow','narrative'];
    const CAT_LABEL = {meta:'メタ',structure:'構造',temporal:'時間',convergence:'収束・拡散',evaluation:'評価・判断',knowledge:'知識・定義',flow:'流れ・構造',narrative:'ナラティブ'};
    TYPES_META.forEach(t => { (byCategory[t.category] = byCategory[t.category] || []).push(t); });
    const parts = [];
    CAT_ORDER.forEach(cat => {
        if (!byCategory[cat]) return;
        parts.push(`<div class="type-category-header">${CAT_LABEL[cat] || cat}</div>`);
        byCategory[cat].forEach(t => {
            const hasForm = !!TYPE_SCHEMAS[t.css_class];
            parts.push(`<div class="type-card" onclick="selectType('${t.css_class}')" title="${t.use_when}">
                <div class="type-name">${t.name}${hasForm ? ' ✓' : ''}</div>
                <div class="type-geom">${t.geometry}</div>
                <div class="type-meaning">${t.meaning}</div>
            </div>`);
        });
    });
    grid.innerHTML = parts.join('');
    openModal('picker-modal');
}

let currentType = null;
let currentData = {};

function selectType(cssClass) {
    closeModal('picker-modal');
    const schema = TYPE_SCHEMAS[cssClass];
    if (!schema) {
        // No form yet — fall back to inserting a minimal snippet
        const sep = editor.value.trim() ? '\\n\\n---\\n\\n' : '';
        const meta = TYPES_META.find(t => t.css_class === cssClass);
        editor.value += sep + `<!-- _class: ${cssClass} -->\\n# ${meta ? meta.meaning : 'タイトル'}\\n`;
        updateStats();
        triggerAutoPreview();
        editor.focus();
        return;
    }
    currentType = cssClass;
    currentData = {};
    // Init defaults
    schema.fields.forEach(f => {
        currentData[f.name] = (f.default !== undefined)
            ? (f.type === 'array' ? JSON.parse(JSON.stringify(f.default)) : f.default)
            : (f.type === 'array' ? [] : (f.type === 'checkbox' ? false : ''));
    });
    document.getElementById('form-title').textContent = schema.label;
    document.getElementById('form-body').innerHTML = buildFormHtml(schema);
    openModal('form-modal');
}

function buildFormHtml(schema) {
    return schema.fields.map(f => buildFieldHtml(f, currentData[f.name], f.name)).join('');
}

function buildFieldHtml(f, value, path) {
    if (f.type === 'text') {
        return `<div class="form-row">
            <label>${f.label}</label>
            <input type="text" value="${escAttr(value||'')}" oninput="setField('${path}', this.value)">
            ${f.hint ? `<div class="hint">${f.hint}</div>` : ''}
        </div>`;
    } else if (f.type === 'textarea') {
        return `<div class="form-row">
            <label>${f.label}</label>
            <textarea oninput="setField('${path}', this.value)">${esc(value||'')}</textarea>
            ${f.hint ? `<div class="hint">${f.hint}</div>` : ''}
        </div>`;
    } else if (f.type === 'checkbox') {
        return `<div class="form-row">
            <label><input type="checkbox" ${value?'checked':''} onchange="setField('${path}', this.checked)"> ${f.label}</label>
        </div>`;
    } else if (f.type === 'array') {
        const items = (value||[]).map((item, i) => buildArrayItemHtml(f, item, path, i)).join('');
        return `<div class="form-row">
            <label>${f.label}</label>
            <div class="array-items" id="arr-${path}">${items}</div>
            <button type="button" class="add-item-btn" onclick="addArrayItem('${path}')">+ 項目を追加</button>
        </div>`;
    }
    return '';
}

function buildArrayItemHtml(f, item, path, index) {
    const sub = f.subfields.map(sf => {
        const val = item[sf.name] || '';
        const subPath = `${path}[${index}].${sf.name}`;
        if (sf.type === 'checkbox') {
            return `<label style="font-size:0.8em"><input type="checkbox" ${item[sf.name]?'checked':''} onchange="setField('${subPath}', this.checked)"> ${sf.label}</label>`;
        }
        return `<div>
            <label style="font-size:0.75em; margin-bottom:2px">${sf.label}</label>
            <input type="text" value="${escAttr(val)}" oninput="setField('${subPath}', this.value)" style="padding:5px 8px">
        </div>`;
    }).join('');
    return `<div class="array-item">
        <div style="display:flex; flex-direction:column; gap:4px; flex:1">${sub}</div>
        <button type="button" class="remove-btn" onclick="removeArrayItem('${path}', ${index})">×</button>
    </div>`;
}

function escAttr(s) { return String(s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;'); }

function setField(path, value) {
    // path like "h1" or "items[0].label"
    const m = path.match(/^(\\w+)(?:\\[(\\d+)\\]\\.(\\w+))?$/);
    if (!m) return;
    if (m[2] !== undefined) {
        currentData[m[1]][+m[2]][m[3]] = value;
    } else {
        currentData[m[1]] = value;
    }
}

function addArrayItem(path) {
    const schema = TYPE_SCHEMAS[currentType];
    const field = schema.fields.find(f => f.name === path);
    const newItem = {};
    field.subfields.forEach(sf => { newItem[sf.name] = sf.default !== undefined ? sf.default : (sf.type === 'checkbox' ? false : ''); });
    currentData[path].push(newItem);
    rerenderArray(path);
}

function removeArrayItem(path, index) {
    currentData[path].splice(index, 1);
    rerenderArray(path);
}

function rerenderArray(path) {
    const schema = TYPE_SCHEMAS[currentType];
    const field = schema.fields.find(f => f.name === path);
    const container = document.getElementById('arr-' + path);
    container.innerHTML = currentData[path].map((item, i) => buildArrayItemHtml(field, item, path, i)).join('');
}

function submitForm() {
    const schema = TYPE_SCHEMAS[currentType];
    const md = schema.toMd(currentData);
    const sep = editor.value.trim() ? '\\n\\n---\\n\\n' : '';
    editor.value += sep + md + '\\n';
    updateStats();
    triggerAutoPreview();
    closeModal('form-modal');
    editor.focus();
    editor.scrollTop = editor.scrollHeight;
}

function insertSnippet(type) {
    if (type === 'plain') selectType('plain');
    else if (type === 'bullets') {
        const sep = editor.value.trim() ? '\\n\\n---\\n\\n' : '';
        editor.value += sep + `# 箇条書き\\n- ポイント1\\n- ポイント2\\n- ポイント3\\n`;
        updateStats(); triggerAutoPreview();
    }
    else if (type === 'divider') selectType('divider');
}

async function loadSample(name) {
    try {
        const r = await fetch('/editor/sample/' + name);
        if (!r.ok) throw new Error(await r.text());
        editor.value = await r.text();
        updateStats();
        editor.scrollTop = 0;
    } catch(e) {
        setStatus('サンプル読込失敗: ' + e.message, 'err');
    }
}

function setStatus(msg, kind) {
    statusEl.textContent = msg;
    statusEl.className = 'status ' + kind;
    if (kind === 'ok') setTimeout(() => { statusEl.className = 'status'; }, 3000);
}

async function generate() {
    const btn = document.getElementById('gen-btn');
    const md = editor.value;
    if (!md.trim()) { setStatus('Markdownが空です', 'err'); return; }
    btn.disabled = true; btn.textContent = '生成中...';
    try {
        const form = new FormData();
        form.append('markdown', md);
        form.append('palette', document.getElementById('palette').value);
        form.append('font_scale', fsRange.value);
        form.append('output_name', document.getElementById('output-name').value || 'slides.pptx');
        const r = await fetch('/editor/generate', { method: 'POST', body: form });
        if (!r.ok) throw new Error(await r.text());
        const blob = await r.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = document.getElementById('output-name').value || 'slides.pptx';
        document.body.appendChild(a); a.click(); a.remove();
        URL.revokeObjectURL(url);
        setStatus('生成完了 → ダウンロード', 'ok');
    } catch(e) {
        setStatus('生成失敗: ' + e.message, 'err');
    } finally {
        btn.disabled = false;
        btn.textContent = '→ PPTX を生成してダウンロード';
    }
}

async function refreshPreview() {
    const panel = document.getElementById('preview-content');
    const md = editor.value;
    if (!md.trim()) {
        panel.innerHTML = '<div class="preview-empty">エディタに内容を入れて<br>「更新」を押してください</div>';
        return;
    }
    panel.innerHTML = '<div class="preview-loading">レンダリング中...<br><small>(LibreOffice経由・数秒かかります)</small></div>';
    try {
        const form = new FormData();
        form.append('markdown', md);
        form.append('palette', document.getElementById('palette').value);
        form.append('font_scale', fsRange.value);
        const r = await fetch('/editor/preview', { method: 'POST', body: form });
        if (!r.ok) throw new Error(await r.text());
        const data = await r.json();
        if (!data.slides || data.slides.length === 0) {
            panel.innerHTML = '<div class="preview-empty">プレビュー生成失敗</div>';
            return;
        }
        panel.innerHTML = data.slides.map((url, i) => `
            <div class="slide-thumb">
                <img src="${url}" alt="slide ${i+1}" loading="lazy">
                <div class="caption">Slide ${i+1}</div>
            </div>
        `).join('');
    } catch(e) {
        panel.innerHTML = '<div class="preview-empty" style="color:#c62828">エラー: ' + e.message + '</div>';
    }
}

// Debounced auto-preview
let previewTimer = null;
let autoPreviewEnabled = true;
function triggerAutoPreview() {
    if (!autoPreviewEnabled) return;
    if (previewTimer) clearTimeout(previewTimer);
    previewTimer = setTimeout(() => refreshPreview(), 1500);
}
editor.addEventListener('input', triggerAutoPreview);

// Initialize
loadTypeMeta();
updateStats();
</script>
</body>
</html>"""


INDEX_HTML = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>marp-pptx Web UI</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, 'Segoe UI', 'Hiragino Sans', sans-serif; background: #f7f7f7; color: #1a1a1a; line-height: 1.6; }
.container { max-width: 900px; margin: 40px auto; padding: 0 20px; }
h1 { font-size: 1.8em; margin-bottom: 8px; }
.subtitle { color: #999; margin-bottom: 32px; }
.card { background: white; border-radius: 8px; padding: 32px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); margin-bottom: 24px; }
label { display: block; font-weight: 600; margin-bottom: 8px; margin-top: 12px; }
select, input[type="file"], input[type="text"] { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; margin-bottom: 16px; }
button { background: #1a1a1a; color: white; border: none; padding: 12px 32px; border-radius: 4px; font-size: 1em; cursor: pointer; margin-right: 8px; }
button:hover { background: #333; }
button.secondary { background: white; color: #1a1a1a; border: 1px solid #ddd; }
.types-table { width: 100%; border-collapse: collapse; font-size: 0.9em; }
.types-table th { background: #1a1a1a; color: white; padding: 10px 12px; text-align: left; }
.types-table td { padding: 8px 12px; border-bottom: 1px solid #eee; }
.types-table tr:nth-child(even) { background: #f9f9f9; }
.cat-badge { display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 0.8em; background: #e8e8e8; }
.tabs { display: flex; border-bottom: 2px solid #ddd; margin-bottom: 20px; }
.tabs a { padding: 10px 20px; color: #666; text-decoration: none; border-bottom: 2px solid transparent; margin-bottom: -2px; }
.tabs a.active { color: #1a1a1a; border-color: #1a1a1a; font-weight: 600; }
</style>
</head>
<body>
<div class="container">
<h1>marp-pptx</h1>
<p class="subtitle">Marp Markdown → Editable PowerPoint (49 semantic slide types)</p>

<div class="tabs">
<a href="/" class="active">変換</a>
<a href="/types-page">型一覧</a>
</div>

<div class="card" style="background:#1a1a1a; color:white;">
<h2 style="margin-bottom:16px">✏️ ブラウザで直接編集 → PPTX 生成</h2>
<p style="margin-bottom:16px; color:#ccc; font-size:0.9em">
.md ファイルを用意せず、その場で Markdown を書いて PPTX にします。型の挿入ボタンあり。
</p>
<a href="/editor"><button type="button" style="background:white; color:#1a1a1a; cursor:pointer">→ エディタを開く</button></a>
</div>

<div class="card">
<h2 style="margin-bottom:16px">① 簡易変換 (設定なしで即ダウンロード)</h2>
<form action="/convert" method="post" enctype="multipart/form-data">
<label>Markdown File (.md)</label>
<input type="file" name="file" accept=".md" required>
<label>Palette</label>
<select name="palette">
<option value="">Default (monochrome)</option>
{% for p in palettes %}<option value="{{ p }}">{{ p }}</option>{% endfor %}
</select>
<button type="submit">→ PPTX に変換してダウンロード</button>
</form>
</div>

<div class="card">
<h2 style="margin-bottom:16px">② 調整画面 (スライド分析 + フォント倍率 + パレット)</h2>
<form action="/preview" method="post" enctype="multipart/form-data">
<label>Markdown File (.md)</label>
<input type="file" name="file" accept=".md" required>
<button type="submit">→ プレビュー画面へ</button>
</form>
</div>
</div>
</body>
</html>"""


TYPES_PAGE_HTML = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>marp-pptx: Slide Types</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, 'Segoe UI', 'Hiragino Sans', sans-serif; background: #f7f7f7; color: #1a1a1a; line-height: 1.6; }
.container { max-width: 1100px; margin: 40px auto; padding: 0 20px; }
h1 { margin-bottom: 20px; }
.tabs { display: flex; border-bottom: 2px solid #ddd; margin-bottom: 20px; }
.tabs a { padding: 10px 20px; color: #666; text-decoration: none; border-bottom: 2px solid transparent; margin-bottom: -2px; }
.tabs a.active { color: #1a1a1a; border-color: #1a1a1a; font-weight: 600; }
table { width: 100%; border-collapse: collapse; font-size: 0.9em; background: white; }
th { background: #1a1a1a; color: white; padding: 10px 12px; text-align: left; }
td { padding: 8px 12px; border-bottom: 1px solid #eee; }
tr:nth-child(even) { background: #f9f9f9; }
.cat-badge { display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 0.8em; background: #e8e8e8; }
</style>
</head>
<body>
<div class="container">
<h1>型一覧 ({{ types|length }})</h1>
<div class="tabs">
<a href="/">変換</a>
<a href="/types-page" class="active">型一覧</a>
</div>
<table>
<thead><tr><th>型</th><th>カテゴリ</th><th>図形</th><th>意味</th><th>使いどころ</th></tr></thead>
<tbody>
{% for t in types %}
<tr>
<td><code>{{ t.name }}</code></td>
<td><span class="cat-badge">{{ categories[t.category] }}</span></td>
<td>{{ t.geometry }}</td>
<td>{{ t.meaning }}</td>
<td>{{ t.use_when }}</td>
</tr>
{% endfor %}
</tbody>
</table>
</div>
</body>
</html>"""


PREVIEW_HTML = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>marp-pptx: Preview & Adjust</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, 'Segoe UI', 'Hiragino Sans', sans-serif; background: #f7f7f7; color: #1a1a1a; line-height: 1.6; }
.layout { display: grid; grid-template-columns: 320px 1fr; min-height: 100vh; }
aside { background: white; border-right: 1px solid #ddd; padding: 24px; position: sticky; top: 0; height: 100vh; overflow-y: auto; }
main { padding: 24px 32px; overflow-y: auto; }
h1 { font-size: 1.4em; margin-bottom: 16px; }
h2 { font-size: 1.1em; margin: 16px 0 8px; color: #555; }
label { display: block; font-weight: 600; margin: 12px 0 6px; font-size: 0.9em; }
select, input { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 0.9em; }
input[type="range"] { padding: 0; }
.slider-row { display: flex; align-items: center; gap: 8px; }
.slider-val { min-width: 36px; font-variant-numeric: tabular-nums; font-size: 0.85em; color: #666; }
button { width: 100%; background: #1a1a1a; color: white; border: none; padding: 14px; border-radius: 4px; font-size: 1em; cursor: pointer; margin-top: 20px; font-weight: 600; }
button:hover { background: #333; }
button.secondary { background: white; color: #1a1a1a; border: 1px solid #ddd; }
.slide-card { background: white; border-radius: 6px; padding: 16px 20px; margin-bottom: 12px; border-left: 4px solid #3d5a80; }
.slide-card.warning { border-left-color: #e07a5f; }
.slide-num { color: #999; font-size: 0.85em; }
.slide-type { display: inline-block; background: #e8e8e8; padding: 2px 8px; border-radius: 3px; font-size: 0.8em; font-family: ui-monospace, monospace; margin-left: 6px; }
.slide-h1 { font-size: 1.1em; font-weight: 600; margin: 6px 0; }
.slide-stats { font-size: 0.85em; color: #666; }
.back-link { color: #666; text-decoration: none; font-size: 0.9em; }
.back-link:hover { color: #1a1a1a; }
</style>
</head>
<body>
<div class="layout">
<aside>
<a href="/" class="back-link">← 戻る</a>
<h1 style="margin-top:12px">設定</h1>
<form action="/generate" method="post" id="gen-form">
<input type="hidden" name="session_id" value="{{ session_id }}">

<h2>パレット</h2>
<label>Color Palette</label>
<select name="palette">
<option value="">Default (monochrome)</option>
{% for p in palettes %}<option value="{{ p }}">{{ p }}</option>{% endfor %}
</select>

<h2>サイズ</h2>
<label>Font Scale (0.7 - 1.3)</label>
<div class="slider-row">
<input type="range" name="font_scale" min="0.7" max="1.3" step="0.05" value="1.0" id="fs-range">
<span class="slider-val" id="fs-val">1.00</span>
</div>

<h2>出力</h2>
<label>Filename</label>
<input type="text" name="output_name" value="{{ filename_base }}_editable.pptx">

<button type="submit">→ PPTX を生成してダウンロード</button>
</form>

<div style="margin-top:20px; padding-top:16px; border-top:1px solid #eee; font-size:0.8em; color:#888;">
<p>※ margin_scale / per-slide override は v0.2 で対応予定 (ROADMAP参照)</p>
</div>
</aside>

<main>
<h1>{{ filename }} — {{ slides|length }} slides</h1>
<p style="color:#666; margin-bottom:20px">
左のパネルで設定を調整 → 下部の「PPTX を生成」ボタンでダウンロード。
</p>

{% for s in slides %}
<div class="slide-card {% if s.warning %}warning{% endif %}">
<span class="slide-num">Slide {{ loop.index }}</span>
<span class="slide-type">{{ s.type_display }}</span>
{% if s.h1 %}<div class="slide-h1">{{ s.h1 }}</div>{% endif %}
<div class="slide-stats">
{% if s.h2 %}<span>H2: {{ s.h2 }}</span> · {% endif %}
<span>{{ s.char_count }} chars</span>
{% if s.bullet_count %} · <span>{{ s.bullet_count }} bullets</span>{% endif %}
{% if s.table_rows %} · <span>{{ s.table_rows }} table rows</span>{% endif %}
{% if s.has_image %} · <span>🖼 image</span>{% endif %}
{% if s.has_math %} · <span>∑ math</span>{% endif %}
</div>
{% if s.warning %}<div style="color:#e07a5f; font-size:0.85em; margin-top:6px">⚠ {{ s.warning }}</div>{% endif %}
</div>
{% endfor %}
</main>
</div>

<script>
const range = document.getElementById('fs-range');
const val = document.getElementById('fs-val');
range.addEventListener('input', () => { val.textContent = parseFloat(range.value).toFixed(2); });
</script>
</body>
</html>"""


# Session-based storage of uploaded MD files
_SESSIONS: dict[str, Path] = {}


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10MB

    def _palettes() -> list[str]:
        palettes_dir = Path(__file__).parent.parent / "data" / "themes" / "palettes"
        return sorted(
            p.stem.replace("academic-", "")
            for p in palettes_dir.glob("academic-*.css")
        )

    @app.route("/")
    def index():
        return render_template_string(INDEX_HTML, palettes=_palettes())

    @app.route("/editor")
    def editor():
        return render_template_string(EDITOR_HTML, palettes=_palettes())

    @app.route("/editor/sample/<name>")
    def editor_sample(name: str):
        """Return a sample Markdown document to populate the editor."""
        templates_dir = Path(__file__).parent.parent / "data" / "templates"
        if name == "minimal":
            md = """---
marp: true
theme: academic
---

<!-- _class: title -->
# 発表タイトル
## サブタイトル
発表者名 / 2026年

---

# 概要
- 背景
- 手法
- 結果
- 考察

---

<!-- _class: end -->
# Thank You
"""
        elif name == "all":
            # Concatenate all 49 templates
            parts = ["---\nmarp: true\ntheme: academic\n---\n"]
            for tpl in sorted(templates_dir.glob("*.md")):
                text = tpl.read_text(encoding="utf-8")
                if text.startswith("---"):
                    end = text.find("---", 3)
                    if end != -1:
                        text = text[end + 3:]
                parts.append(text.strip())
            md = "\n\n---\n\n".join(parts)
        elif name == "academic":
            md = """---
marp: true
theme: academic
---

<!-- _class: title -->
# 研究タイトル
## サブタイトル
山田 太郎 / 2026年4月

---

<!-- _class: agenda -->
# 本日の内容
<div class="agenda-list">
1. 背景と研究目的
2. 提案手法
3. 実験結果
4. 考察とまとめ
</div>

---

<!-- _class: rq -->
# 研究課題
<div class="rq-main">既存手法は大規模データに対してスケールするか？</div>
<div class="rq-sub">計算量 $O(n^2)$ がボトルネックとなっている。</div>

---

<!-- _class: sandwich -->
# 提案手法
<div class="top">
<div class="lead">従来の $O(n^2)$ を $O(n \\log n)$ に改善。</div>
</div>
<div class="columns">
<div>

### 従来
- 計算量: `O(n^2)`
- メモリ: 多い
</div>
<div>

### 提案
- 計算量: `O(n \\log n)`
- メモリ: 少ない
</div>
</div>
<div class="bottom">
<div class="conclusion"><strong>結論:</strong> 大規模データでも実用的な速度を実現。</div>
</div>

---

<!-- _class: kpi -->
# 実験結果
<div class="kpi-container">
<div><span class="kpi-value">97%</span><span class="kpi-label">精度</span></div>
<div><span class="kpi-value">10x</span><span class="kpi-label">高速化</span></div>
<div><span class="kpi-value">50%</span><span class="kpi-label">省メモリ</span></div>
</div>

---

<!-- _class: takeaway -->
# Takeaway
<div class="ta-main">型を選ぶだけで、伝わるプレゼンに</div>
<div class="ta-points">
<ul>
<li>計算量の改善により大規模データに対応</li>
<li>精度は従来と同等</li>
<li>OSS として公開予定</li>
</ul>
</div>

---

<!-- _class: end -->
# Thank You
"""
        else:
            return "unknown sample", 404
        from flask import Response
        return Response(md, mimetype="text/plain; charset=utf-8")

    @app.route("/editor/preview", methods=["POST"])
    def editor_preview():
        """Render MD → PPTX → per-slide PNGs; return list of URLs."""
        from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
        from marp_pptx.parser import parse_marp
        from marp_pptx.builder import PptxBuilder

        md_text = request.form.get("markdown", "")
        if not md_text.strip():
            return jsonify({"slides": []})
        palette_name = request.form.get("palette", "")
        try:
            font_scale = float(request.form.get("font_scale", 1.0))
        except ValueError:
            font_scale = 1.0

        # Cache key based on content + settings
        key_src = f"{md_text}|{palette_name}|{font_scale}".encode("utf-8")
        key = hashlib.md5(key_src).hexdigest()
        out_dir = _PREVIEW_CACHE_DIR / key
        if out_dir.exists():
            pngs = sorted(out_dir.glob("slide-*.png"))
            if pngs:
                return jsonify({"slides": [f"/editor/preview-img/{key}/{p.name}" for p in pngs]})
        out_dir.mkdir(parents=True, exist_ok=True)

        # Build PPTX
        md_path = out_dir / "slides.md"
        md_path.write_text(md_text, encoding="utf-8")
        tc = ThemeConfig.from_css(get_default_theme_path())
        tc.font_scale = max(0.5, min(2.0, font_scale))
        if palette_name:
            pp = get_palette_path(palette_name)
            if pp:
                tc.apply_palette(pp)
        slides = parse_marp(str(md_path))
        builder = PptxBuilder(base_path=out_dir, theme=tc)
        builder.build_all(slides)
        pptx_path = out_dir / "slides.pptx"
        builder.save(str(pptx_path))

        # Render to PNGs
        pngs = _render_pptx_to_pngs(pptx_path, out_dir, dpi=90)
        if not pngs:
            return jsonify({"error": "LibreOffice not available or render failed", "slides": []}), 500
        return jsonify({"slides": [f"/editor/preview-img/{key}/{p.name}" for p in pngs]})

    @app.route("/editor/preview-img/<key>/<name>")
    def editor_preview_img(key: str, name: str):
        """Serve a cached preview PNG."""
        if not key.isalnum() or not name.startswith("slide-") or not name.endswith(".png"):
            return "bad path", 400
        png = _PREVIEW_CACHE_DIR / key / name
        if not png.exists():
            return "not found", 404
        return send_file(str(png), mimetype="image/png")

    @app.route("/editor/generate", methods=["POST"])
    def editor_generate():
        """Generate PPTX from raw Markdown text (no file upload)."""
        md_text = request.form.get("markdown", "")
        if not md_text.strip():
            return "empty markdown", 400
        palette_name = request.form.get("palette", "")
        try:
            font_scale = float(request.form.get("font_scale", 1.0))
        except ValueError:
            font_scale = 1.0
        output_name = request.form.get("output_name") or "slides.pptx"

        tmpdir = Path(tempfile.mkdtemp(prefix="marp_editor_"))
        md_path = tmpdir / "slides.md"
        md_path.write_text(md_text, encoding="utf-8")
        return _do_convert(
            md_path=md_path,
            palette_name=palette_name,
            font_scale=font_scale,
            output_name=output_name,
        )

    @app.route("/types-page")
    def types_page():
        from marp_pptx.types import TYPE_REGISTRY, CATEGORIES
        return render_template_string(
            TYPES_PAGE_HTML,
            types=TYPE_REGISTRY,
            categories=CATEGORIES,
        )

    @app.route("/convert", methods=["POST"])
    def convert():
        return _do_convert(
            palette_name=request.form.get("palette", ""),
            font_scale=1.0,
            output_name=None,
        )

    @app.route("/preview", methods=["POST"])
    def preview():
        from marp_pptx.parser import parse_marp
        from marp_pptx.types import get_type_info

        f = request.files.get("file")
        if not f:
            return "No file uploaded", 400

        # Save to session
        session_id = uuid.uuid4().hex
        tmpdir = Path(tempfile.mkdtemp(prefix="marp_preview_"))
        md_path = tmpdir / (f.filename or "slides.md")
        f.save(str(md_path))
        _SESSIONS[session_id] = md_path

        slides_data = parse_marp(str(md_path))
        slides = []
        for sd in slides_data:
            info = get_type_info(sd.slide_class) if sd.slide_class else None
            type_display = sd.slide_class or "default"
            char_count = len(sd.raw)
            bullet_count = sum(
                1 for line in sd.body_lines
                if line.strip().startswith(("- ", "* "))
            )
            table_rows = len(sd.table_rows)
            has_image = bool(sd.image_path) or bool(sd.annotation_figure) or bool(sd.result_figure) or bool(sd.gallery_items)
            has_math = bool(sd.eq_main) or bool(sd.eq_system) or "$" in sd.raw
            warning = None
            if sd.slide_class and not info and sd.slide_class not in (
                "cols-2-wide-l", "cols-2-wide-r",
            ):
                warning = f"未知の型: {sd.slide_class}"
            slides.append({
                "type_display": type_display,
                "h1": sd.h1,
                "h2": sd.h2,
                "char_count": char_count,
                "bullet_count": bullet_count,
                "table_rows": table_rows,
                "has_image": has_image,
                "has_math": has_math,
                "warning": warning,
            })

        filename = md_path.name
        filename_base = md_path.stem

        return render_template_string(
            PREVIEW_HTML,
            slides=slides,
            palettes=_palettes(),
            session_id=session_id,
            filename=filename,
            filename_base=filename_base,
        )

    @app.route("/generate", methods=["POST"])
    def generate():
        session_id = request.form.get("session_id", "")
        md_path = _SESSIONS.get(session_id)
        if md_path is None or not md_path.exists():
            return "Session expired. Please re-upload.", 400

        return _do_convert(
            md_path=md_path,
            palette_name=request.form.get("palette", ""),
            font_scale=float(request.form.get("font_scale", 1.0)),
            output_name=request.form.get("output_name") or None,
        )

    def _do_convert(md_path=None, palette_name="", font_scale=1.0, output_name=None):
        from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
        from marp_pptx.parser import parse_marp
        from marp_pptx.builder import PptxBuilder

        if md_path is None:
            f = request.files.get("file")
            if not f:
                return "No file uploaded", 400
            tmpdir = Path(tempfile.mkdtemp())
            md_path = tmpdir / (f.filename or "slides.md")
            f.save(str(md_path))

        tc = ThemeConfig.from_css(get_default_theme_path())
        tc.font_scale = max(0.5, min(2.0, font_scale))
        if palette_name:
            pp = get_palette_path(palette_name)
            if pp:
                tc.apply_palette(pp)

        slides = parse_marp(str(md_path))
        builder = PptxBuilder(base_path=md_path.parent, theme=tc)
        builder.build_all(slides)

        out_name = output_name or (md_path.stem + "_editable.pptx")
        out_path = md_path.parent / out_name
        builder.save(str(out_path))

        return send_file(
            str(out_path),
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    @app.route("/api/types")
    def api_types():
        from marp_pptx.types import TYPE_REGISTRY, CATEGORIES
        data = [
            {
                "name": t.name,
                "css_class": t.css_class,
                "category": t.category,
                "category_ja": CATEGORIES.get(t.category, t.category),
                "geometry": t.geometry,
                "meaning": t.meaning,
                "use_when": t.use_when,
            }
            for t in TYPE_REGISTRY
        ]
        return jsonify(data)

    return app
