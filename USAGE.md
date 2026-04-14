# marp-pptx 使い方ガイド (AI向け)

このドキュメントは、AI エージェントが marp-pptx を使ってユーザーのプレゼン資料を作成するための実用リファレンスです。

## プロジェクトの位置づけ

**このツールは Markdown → PPTX の一方向変換ツールです。** 逆方向（PPTX → MD）は扱いません。

### 設計思想

PPTX で実現できる洗練されたレイアウト（KPI ダッシュボード、ファネル図、2x2 マトリクス、
タイムライン等）を、**Markdown の記述だけで再現する**ための試みです。PowerPoint の表現力の
良さを、Markdown の編集しやすさと組み合わせます。

- **MD → PPTX**: 本ツールが担当（この方向の品質にフォーカス）
- **PPTX → MD**: 非対応（PPTX の見た目を Markdown に自動逆変換することはできない）
- **PPTX を直接編集**: 出力された PPTX は完全編集可能。仕上げは PowerPoint で行う想定

### 型ライブラリ = PPTX 風サンプルの MD 実装集

`templates/` と `src/marp_pptx/data/templates/` にある 49 種類の MD ファイルは、
「PPTX で見栄え良く表現される各種スライドパターンを、どう Markdown で書けば再現できるか」
を調査・実装したサンプル集です。AI がユーザーのプレゼンを組む時は、これらをテンプレートとして参照してください。

## TL;DR

```bash
pip install -e .                                    # インストール
marp-pptx convert slides.md -o out.pptx             # 変換
marp-pptx convert slides.md -o out.pptx -p navy     # パレット指定
marp-pptx convert slides.md --font-scale 1.15       # フォント拡大
marp-pptx types                                     # 型一覧
marp-pptx preview -o catalog.pptx                   # 全型のビジュアル例
marp-pptx serve --port 8080                         # Web UI (プレビュー調整画面あり)
```

## Markdown の書き方と PPTX への対応

### 記法対応表 (MD → PPTX)

| Markdown 記法 | PPTX での表現 | 備考 |
|---|---|---|
| `# 見出し` | スライドH1 (大見出し) | 自動で左バー装飾 |
| `## 見出し` | H2 (サブ見出し) | 2次色で表示 |
| `### 見出し` | H3 | ミュート色 |
| `**太字**` | `<b>` (bold run) | bullet内でも有効 |
| `*斜体*` | **非対応**（デザイン上あえて無効化） | 太字で代替推奨 |
| `` `code` `` | モノスペースフォント | `code` 型スライドで推奨 |
| `- 項目` / `* 項目` | 箇条書き (bullet) | `•` マーカー自動付与 |
| `1. 項目` | 番号付きリスト | |
| `[text](url)` | ハイパーリンク | PPTX 上で機能する |
| `![w:800](img.png)` | 画像挿入 | `w:N` で幅指定 (px) |
| `$x^2$` | インライン数式 (OMML) | PowerPoint 上で編集可 |
| `$$\frac{a}{b}$$` | ディスプレイ数式 | 中央配置 |
| `\| A \| B \|` | 表 | 区切り行 `\|---\|---\|` 必須 |
| `> 引用` | **非対応**（`quote` 型で代替） | `<!-- _class: quote -->` を使う |
| ソフトラップ (改行のみ) | 同じ段落に統合 | 可読性のため改行しても段落は分かれない |
| 空行 | 新しい段落 | 段落を分けたい時は空行を入れる |

### 改行と段落のルール（重要）

Markdown 標準に準拠：
- **空行で段落が分かれる**（`<p>` が切り替わる）
- **改行だけ**では段落は分かれない（ソフトラップ・同じ段落として結合）

```markdown
これは1つの段落で
可読性のため
改行しています。

これは次の段落です。
```

→ PPTX 上では「これは1つの段落で 可読性のため 改行しています。」が1段落、
「これは次の段落です。」が別段落。同一テキストボックス内、別パラグラフ。

### 強調の優先度

1. **最優先**: `**bold**` — シンプルで確実
2. HTML の `<strong>` — `strip_html` で消えるので非推奨
3. 斜体 `*text*` — デザイン意図で無効化済み

### 日本語フォント

自動対応しています。CSS `--font-ea` で指定したフォントが全ての `<a:ea>` 属性に注入され、
英数字は `--font-body` が適用されます。混在文は同一テキストボックス内で自然に描画されます。

## 基本原則: 「型」を選んで書く

このツールの核心は **49種類のセマンティックなスライド型**。
ユーザーが「何を伝えたいか」を聞いたら、まず **どの型を使うか** を決める。

### 型選択の思考フロー

| ユーザーの意図 | 選ぶべき型 |
|---|---|
| 始まり | `title` |
| 章の区切り | `divider` |
| 予定・目次 | `agenda` |
| 2つを並列比較 | `cols-2` |
| 3つを分類 | `cols-3` |
| 概要→詳細→結論 | `sandwich` |
| 賛否を示す | `pros-cons` |
| 2軸で評価 | `zone-matrix` |
| 時間の流れ（横） | `timeline-h` |
| 時間の流れ（縦） | `timeline` |
| 手順 | `steps` |
| ビフォーアフター | `before-after` |
| 絞り込み（多→少） | `funnel` |
| 積み重ね | `stack` |
| 数値・KPI | `kpi` |
| 複数の結果 | `multi-result` |
| 単一の結果＋分析 | `result` |
| 2つの結果並列 | `result-dual` |
| 用語の定義 | `definition` |
| 1つの数式 | `equation` |
| 連立式・最適化問題 | `equations` |
| 図＋キャプション | `figure` |
| 図＋注釈 | `annotation` |
| 構造図 | `diagram` |
| 複数画像 | `gallery-img` |
| 横長画像で没入感 | `panorama` |
| コード | `code` |
| 表 | `table-slide` |
| プロセス＋詳細 | `zone-process` |
| フロー (A→B→C) | `zone-flow` |
| 2項比較 (VS) | `zone-compare` |
| チェックリスト | `checklist` |
| 引用 | `quote` |
| 沿革 | `history` |
| 人物紹介 | `profile` |
| 全体像 | `overview` |
| 強調 (1つだけ) | `highlight` |
| カード状一覧 | `card-grid` |
| 左右分割テキスト | `split-text` |
| 研究質問 | `rq` |
| まとめ | `summary` |
| キーメッセージ | `takeaway` |
| 参考文献 | `references` |
| 補足 | `appendix` |
| 終わり | `end` |

## Markdown の書き方

### 基本構造

```markdown
---
marp: true
theme: academic
math: katex
---

<!-- _class: title -->
# プレゼンタイトル
## サブタイトル
発表者名 / 2026-04

---

<!-- _class: agenda -->
# 本日の内容
<div class="agenda-list">
1. 背景
2. 手法
3. 結果
4. まとめ
</div>

---

<!-- _class: end -->
# Thank You
```

- スライド区切りは `---`
- 型の指定は `<!-- _class: 型名 -->`
- フロントマター (`---...---`) は1つだけ先頭に
- 各型は**特定のHTML構造**を期待する（下記参照）

### 各型のテンプレート

#### title — 表紙

```markdown
<!-- _class: title -->
# メインタイトル
## サブタイトル
発表者: 山田太郎
2026年4月14日
```

#### divider — 章区切り

```markdown
<!-- _class: divider -->
# 第2章
## 提案手法
```

#### cols-2 / cols-3 — 並列・分類

```markdown
<!-- _class: cols-2 -->
# 比較
<div class="columns">
<div>
### 従来手法
- 遅い
- メモリ多い
</div>
<div>
### 提案手法
- 高速
- 省メモリ
</div>
</div>
```

#### sandwich — 概要→詳細→結論

```markdown
<!-- _class: sandwich -->
# タイトル
<div class="top">
<div class="lead">リード文（全体を要約する1行）</div>
</div>
<div class="columns">
<div>詳細1</div>
<div>詳細2</div>
</div>
<div class="bottom">
<div class="conclusion"><strong>結論：</strong>...</div>
</div>
```

#### equation — 単一数式

```markdown
<!-- _class: equation -->
# ベイズの定理
<div class="eq-main">
$$P(A|B) = \frac{P(B|A)P(A)}{P(B)}$$
</div>
<div class="eq-desc">
<span>$P(A|B)$</span><span>事後確率</span>
<span>$P(B|A)$</span><span>尤度</span>
<span>$P(A)$</span><span>事前確率</span>
</div>
```

#### equations — 連立式・最適化問題

```markdown
<!-- _class: equations -->
# 最適化問題
<div class="eq-system">
<div class="row"><span class="label">minimize</span> $$f(x) = \|Ax - b\|^2$$</div>
<div class="row"><span class="label">subject to</span> $$Ax \le b$$</div>
<div class="row"><span class="label"></span> $$x \ge 0$$</div>
</div>
```

#### figure — 図＋キャプション

```markdown
<!-- _class: figure -->
# 実験環境
![w:800](assets/setup.png)
<div class="caption"><span class="fig-num">Fig. 1</span> 装置の概観</div>
```

#### timeline-h / timeline — 時系列

```markdown
<!-- _class: timeline-h -->
# プロジェクト進行
<div class="tl-h-container">
<div class="tl-h-item">
<div><span class="tl-h-year">2024</span><span class="tl-h-text">企画</span></div>
</div>
<div class="tl-h-item highlight">
<div><span class="tl-h-year">2025</span><span class="tl-h-text">開発</span></div>
</div>
<div class="tl-h-item">
<div><span class="tl-h-year">2026</span><span class="tl-h-text">リリース</span></div>
</div>
</div>
```

`highlight` クラスを付けると強調色になる。

#### steps — 手順

```markdown
<!-- _class: steps -->
# 使い方
<div class="st-container">
<div><span class="st-num">1</span><span class="st-title">インストール</span><span class="st-body">pip install で導入</span></div>
<div><span class="st-num">2</span><span class="st-title">設定</span><span class="st-body">config.yaml を編集</span></div>
<div><span class="st-num">3</span><span class="st-title">実行</span><span class="st-body">run コマンドで起動</span></div>
</div>
```

#### kpi — 数値強調

```markdown
<!-- _class: kpi -->
# 成果
<div class="kpi-container">
<div><span class="kpi-value">98%</span><span class="kpi-label">精度</span></div>
<div><span class="kpi-value">1.2s</span><span class="kpi-label">推論時間</span></div>
<div><span class="kpi-value">10x</span><span class="kpi-label">高速化</span></div>
</div>
```

#### pros-cons — 賛否

```markdown
<!-- _class: pros-cons -->
# 提案手法の評価
<div class="pc-pros">
<ul><li>高速</li><li>省メモリ</li></ul>
</div>
<div class="pc-cons">
<ul><li>実装コストが高い</li><li>依存関係が多い</li></ul>
</div>
```

#### zone-flow — フロー

```markdown
<!-- _class: zone-flow -->
# 処理フロー
<div class="zf-container">
<div><span class="zf-label">入力</span><span class="zf-body">画像 (224x224)</span></div>
<div><span class="zf-label">特徴抽出</span><span class="zf-body">ResNet50</span></div>
<div><span class="zf-label">分類</span><span class="zf-body">FC層</span></div>
</div>
```

矢印は自動で挿入される。

#### zone-matrix — 2x2評価

```markdown
<!-- _class: zone-matrix -->
# 重要度×緊急度
<div class="zm-xlabel">重要度</div>
<div class="zm-ylabel">緊急度</div>
<div class="zm-tl"><span class="zm-label">高緊急・低重要</span><span class="zm-body">委譲</span></div>
<div class="zm-tr"><span class="zm-label">高緊急・高重要</span><span class="zm-body">即実行</span></div>
<div class="zm-bl"><span class="zm-label">低緊急・低重要</span><span class="zm-body">削除</span></div>
<div class="zm-br"><span class="zm-label">低緊急・高重要</span><span class="zm-body">計画</span></div>
```

#### funnel — 絞り込み

```markdown
<!-- _class: funnel -->
# 採用プロセス
<div class="fn-container">
<div><span class="fn-label">応募</span><span class="fn-value">1,000人</span></div>
<div><span class="fn-label">書類通過</span><span class="fn-value">200人</span></div>
<div><span class="fn-label">面接通過</span><span class="fn-value">50人</span></div>
<div><span class="fn-label">採用</span><span class="fn-value">10人</span></div>
</div>
```

#### before-after — 変化

```markdown
<!-- _class: before-after -->
# 改善結果
<div class="ba-before">
<span class="ba-label">Before</span>
<span class="ba-body">処理時間 5秒</span>
</div>
<div class="ba-after">
<span class="ba-label">After</span>
<span class="ba-body">処理時間 0.5秒</span>
</div>
```

#### quote — 引用

```markdown
<!-- _class: quote -->
# 引用
<div class="qt-text">
プログラムは人間が読むために書くべきであり、
たまたま機械が実行できるに過ぎない。
</div>
<div class="qt-source">Harold Abelson</div>
```

#### definition — 定義

```markdown
<!-- _class: definition -->
# 定義
<div class="df-term">機械学習</div>
<div class="df-body">明示的にプログラムされることなく、データから学習してタスクを実行する能力をコンピュータに与える研究分野。</div>
<div class="df-note">Arthur Samuel (1959)</div>
```

#### code — コード

````markdown
<!-- _class: code -->
# 実装例
<div class="cd-code">
```python
def fibonacci(n):
    if n < 2:
        return n
    return fibonacci(n-1) + fibonacci(n-2)
```
</div>
<div class="cd-desc">再帰によるフィボナッチ数列の実装</div>
````

#### table-slide — 表

```markdown
<!-- _class: table-slide -->
# 比較表
| 手法 | 精度 | 速度 |
|------|-----:|-----:|
| A    | 85%  | 1.0s |
| B    | 92%  | 1.5s |
| **Ours** | **97%** | **0.8s** |
```

#### takeaway — キーメッセージ

```markdown
<!-- _class: takeaway -->
# Takeaway
<div class="ta-main">型を選ぶだけで、伝わるプレゼンになる</div>
<div class="ta-points">
<ul>
<li>49種類の意味的な型</li>
<li>PPTXとして編集可能</li>
<li>日本語・数式対応</li>
</ul>
</div>
```

#### summary — まとめ

```markdown
<!-- _class: summary -->
# まとめ
<ol class="summary-points">
<li>提案手法は従来比10倍高速</li>
<li>精度は同等を維持</li>
<li>実装はOSSとして公開</li>
</ol>
```

#### references — 参考文献

```markdown
<!-- _class: references -->
# 参考文献
<ol>
<li><span class="author">Smith et al.</span> <span class="title">Fast Methods.</span> <span class="venue">NeurIPS 2024.</span></li>
<li><span class="author">Yamada</span> <span class="title">機械学習入門.</span> <span class="venue">Ohmsha, 2023.</span></li>
</ol>
```

#### end — 終わり

```markdown
<!-- _class: end -->
# Thank You
Questions?
```

## パレット（配色）

10種類用意されている。ユーザーの雰囲気に合わせて選ぶ：

| パレット | 雰囲気 |
|---|---|
| `mono` | モノクロ（標準・学術） |
| `navy` | 紺・信頼感 |
| `copper` | 銅・温かみ |
| `earth` | 大地・自然 |
| `forest` | 深緑・落ち着き |
| `ink` | 墨・和風 |
| `ocean` | 青・爽やか |
| `slate` | スレート・ビジネス |
| `violet` | 紫・創造的 |
| `wine` | ワイン・高級感 |

```bash
marp-pptx convert slides.md -p navy
```

## AI が資料作成を依頼されたときの手順

1. **ユーザーの目的を聞く**：何を、誰に、どう伝えたいか
2. **構成を型で設計**：各スライドに型を割り当てる
   - 例：`title → agenda → rq → figure → result → pros-cons → summary → takeaway → end`
3. **Markdown を書く**：上記テンプレート通りに型のHTML構造を埋める
4. **変換**：`marp-pptx convert` で PPTX 生成
5. **確認**：スライド数が合っているか、画像が含まれているか

## よくあるハマりどころ

- **型の指定を忘れると** `default` 型（単なる箇条書き）になる
- **HTML構造を間違えると** 該当箇所が空になる（例：`<div class="kpi-container">` の中に `<div><span class="kpi-value">...` の入れ子が必要）
- **画像パス**：MDファイルからの相対パス
- **フロントマター**は必ず先頭のみ。各スライドに `---` 区切りを入れても frontmatter にならない
- **数式** `$$...$$` は display、`$...$` は inline。OMML変換には Pandoc 必要（なければ matplotlib PNG fallback）
- **日本語フォント**：CSS の `--font-ea` で指定したフォントが自動適用される
- **テンプレート例**は `src/marp_pptx/data/templates/` の 49 ファイルに実例あり

## プログラムから使う（Python API）

```python
from pathlib import Path
from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
from marp_pptx.parser import parse_marp
from marp_pptx.builder import PptxBuilder

tc = ThemeConfig.from_css(get_default_theme_path())
tc.apply_palette(get_palette_path("navy"))

slides = parse_marp("input.md")
builder = PptxBuilder(base_path=Path("."), theme=tc)
builder.build_all(slides)
builder.save("output.pptx")
```

## 型の意味を聞くコード

```python
from marp_pptx.types import TYPE_REGISTRY, get_type_info

info = get_type_info("funnel")
print(info.meaning)   # "絞り込み・選別"
print(info.use_when)  # "多→少の過程を見せるとき"
```

## 依存関係

**必須**：
- Python 3.10+
- `python-pptx`, `lxml`, `Pillow`, `matplotlib`, `click`, `pyyaml`

**推奨**：
- Pandoc（数式を編集可能な OMML にする。なければ matplotlib PNG）

**任意**：
- `pip install marp-pptx[web]` で Flask Web UI

## トラブルシュート

| 症状 | 対処 |
|---|---|
| `pandoc not found` | `brew install pandoc` / `apt install pandoc`（数式PNGにフォールバックするだけなので無視も可） |
| 日本語が豆腐になる | CSS `--font-ea` を変更 or フォントをシステムにインストール |
| 画像が出ない | MDファイルからの相対パスが正しいか確認 |
| スライドが想定数と違う | `---` 区切りを確認（`\n---\n` の前後に空行） |
