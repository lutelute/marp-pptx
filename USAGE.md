# marp-pptx 使い方ガイド (AI向け)

このドキュメントは、AI エージェントが marp-pptx を使ってユーザーのプレゼン資料を作成するための実用リファレンスです。

## プロジェクトの位置づけ

**このツールの主軸は Markdown → PPTX の変換です。** 逆方向（PPTX → MD）は学習データ
作成用のベストエフォート機能として Web UI に用意されています（完全な見た目の復元は保証しません）。

### 設計思想

PPTX で実現できる洗練されたレイアウト（KPI ダッシュボード、ファネル図、2x2 マトリクス、
タイムライン等）を、**Markdown の記述だけで再現する**ための試みです。PowerPoint の表現力の
良さを、Markdown の編集しやすさと組み合わせます。

- **MD → PPTX**: 本ツールの主軸（この方向の品質にフォーカス）
- **PPTX → MD**: ベストエフォート（Web UI の「PPTX を読み込み」。テキスト・構造を抽出するが
  見た目の完全復元は対象外。`marp_pptx.pptx2md` モジュール）
- **PPTX を直接編集**: 出力された PPTX は完全編集可能。仕上げは PowerPoint で行う想定

### 型ライブラリ = PPTX 風サンプルの MD 実装集

`src/marp_pptx/data/templates/` にある 52 種類の MD ファイル（型ごとに 1 つ）は、
「PPTX で見栄え良く表現される各種スライドパターンを、どう Markdown で書けば再現できるか」
を調査・実装したサンプル集です。AI がユーザーのプレゼンを組む時は、これらをテンプレートとして参照してください。
（リポジトリ直下の `templates/` は旧コピーで、新型 50–52 を含みません。正は `src/marp_pptx/data/templates/`。）

## TL;DR

```bash
pip install -e .                                    # インストール（Web UI は pip install -e ".[web]"）
marp-pptx convert slides.md -o out.pptx             # 変換（無指定で claude テーマ）
marp-pptx convert slides.md -o out.pptx -p navy     # パレット指定
marp-pptx convert slides.md --math png              # 数式を画像で焼く（LibreOffice/Keynote 用）
marp-pptx convert slides.md --density keynote       # 投影向けに大きめ・余白広め
marp-pptx convert slides.md --font-scale 1.15       # フォント拡大（0.7–1.3）
marp-pptx doctor slides.md                          # 実測検査（溢れ・重なり・コントラスト・破損）
marp-pptx doctor out.pptx --json --strict           # 機械可読 / warn 以上で exit 1
marp-pptx types                                     # 型一覧（-c でカテゴリ絞り / --json）
marp-pptx themes                                    # テーマ・パレット一覧
marp-pptx preview -o catalog.pptx                   # 全型のビジュアル例（カタログ PPTX）
marp-pptx render-gallery                            # 全型のサムネ PNG を再生成（Web UI ギャラリー用）
marp-pptx serve --port 8080                         # Web UI（フォーム編集＋ライブプレビュー）
```

### convert の主なオプション

| オプション | 既定 | 説明 |
|---|---|---|
| `-o, --output` | `<入力>_editable.pptx` | 出力先 |
| `-p, --palette` | `claude` | パレット名（`minimal` / `navy` / `copper` …。`marp-pptx themes` 参照） |
| `-t, --theme` | — | 任意のパレット CSS パスを直接指定 |
| `--math` | `omml` | `omml`=PowerPoint で編集可（要 pandoc）/ `png`=matplotlib 画像（LibreOffice・Keynote 用） |
| `--density` | `academic` | `academic`=高密度 / `keynote`=投影向けに大きめ（font×1.22・余白×1.12） |
| `--font-scale` | `1.0` | フォント倍率（0.7–1.3） |

## 精度検査（doctor）

```bash
marp-pptx doctor slides.md            # .md はビルドしてから検査。.pptx も直接渡せる
marp-pptx doctor slides.md --json     # findings を JSON で
marp-pptx doctor slides.md --strict   # warn 以上で exit 1（CI）
marp-pptx doctor slides.md -p navy    # ビルドに使うパレット
```

各テキストボックスを、そのボックスが指定する**書体の実際の字送り**で測る。
折り返しは PowerPoint と同じ規則（Latin は語境界、日本語は禁則処理）で再現する。
レンダラ不要・API キー不要・数十 ms。

| kind | severity | 内容 |
|---|---|---|
| `overflow` | error / warn | 折り返し後の高さがボックスを超える（error=切れる、warn=箱が伸びて下を押す）。`word_wrap` 無効なら横はみ出しも |
| `overlap` | error / warn | テキスト同士の重なり（in² 表示）／0.12in 未満の近接 |
| `offslide` | error / warn | スライド外／安全余白 0.35in 内。テーマの帯・フッターバーは自動で除外 |
| `contrast` | warn | 背後の面（カード塗り or スライド背景）に対する WCAG AA 比 |
| `font` | warn / info | 未インストール（計測が近似になる）／プレビューで幅が変わる書体 |
| `package` | error / warn / info | 関係参照切れ・content-type 欠落・拡張子と中身の不一致・未使用メディア |
| `deck` | info | 同一レイアウトが 4 枚以上連続／図表の無いスライド |

**`overflow` は文字を減らして直すのが第一手**（フォント縮小は読みにくさに直結する）。

計測エンジン（`marp_pptx.metrics`）はビルダー・`visuallint` の衝突検出・doctor で共用。
予約したボックスと必要な高さが常に同じ物差しで測られるので、両者が食い違わない。

```python
from marp_pptx import metrics as M
M.measure_pt("見出し", "Helvetica Neue", 30, ea_font="Hiragino Sans")   # 描画幅(pt)
M.wrap_text(text, "Helvetica Neue", 18, 400, ea_font="Hiragino Sans")  # 実際の折り返し
M.fit_size(text, "Helvetica Neue", 400, 60, max_size=30)               # 収まる最大サイズ
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

このツールの核心は **52種類のセマンティックなスライド型**。
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
| 1文を大きく言い切る／論点転換 | `statement` |
| 1つの数字を主役に | `big-number` |
| 数値データをグラフで（編集可） | `chart` |
| 参考文献 | `references` |
| 補足 | `appendix` |
| 終わり | `end` |

**バリエーション型**（親型の `_class` を流用した応用レシピ）：

| ユーザーの意図 | 選ぶべき型 | ベース |
|---|---|---|
| 共通設定＋3条件＋考察を1枚に | `sandwich-3col` | sandwich |
| 数式の各記号を注釈付きで解説 | `equation-annotated` | equation |
| 数式の特定項を色で強調 | `equation-highlight` | equation |
| 図と解説を左右に並べる | `figure-cols` | cols-2 |

## Markdown の書き方

### 基本構造

```markdown
---
marp: true
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
<!-- note: 各章の所要時間に触れる -->

---

<!-- _class: end -->
# Thank You
```

- スライド区切りは `---`
- 型の指定は `<!-- _class: 型名 -->`
- フロントマター (`---...---`) は1つだけ先頭に。`marp: true` だけで十分
  （`theme:` / `math:` フィールドは parser が読みません。**テーマ／パレットは CLI の `-p` で選ぶ**）
- 発表者ノートは `<!-- note: ... -->`（PPTX のノート欄に入る。後述）
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
<li>52種類の意味的な型</li>
<li>PPTXとして編集可能</li>
<li>日本語・数式対応</li>
</ul>
</div>
```

#### statement — 断言・論点転換（全画面1文）

```markdown
<!-- _class: statement -->

本研究は、スパース注意機構に初めて理論的保証を与える。
```

見出し（`#`）も HTML 構造も不要。1 文を全画面の中央に大きく置く。論点の転換や
キーメッセージの「タメ」に使う。

#### big-number — 単一指標の強調

```markdown
<!-- _class: big-number -->
<!-- source: 社内ベンチマーク 2026 (n=1000) -->
# 主要成果
<div class="big-number">
  <span class="bn-value">89.4%</span>
  <span class="bn-label">分類精度</span>
  <span class="bn-caption">従来手法から +4.2pt 改善</span>
</div>
```

ひとつの数字を主役にする。`<!-- source: ... -->` を書くと出典フットノートが付く。

#### chart — データ可視化（編集可能なグラフ）

```markdown
<!-- _class: chart -->
<!-- _chart: column -->
# 系列長ごとの計算時間（相対）
| 系列長 | Transformer | Ours |
|---|---|---|
| 1024 | 1.0 | 0.6 |
| 4096 | 4.2 | 1.1 |
| 16384 | 18.5 | 2.3 |
<div class="chart-caption">同一ハードウェアで測定。</div>
```

表をそのまま**ネイティブの編集可能なグラフ**に変換する（PowerPoint 上で系列・数値を編集可）。
`<!-- _chart: column -->` でグラフ種別を指定（`column` 縦棒 / `bar` 横棒 / `line` 折れ線）。

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

## テーマとパレット（配色）

`-p` で指定する。**テーマ**＝レイアウト＋配色、**パレット**＝配色だけの差し替え。
無指定なら `claude`。一覧は `marp-pptx themes`。

### 主テーマ（レイアウト＋配色）

| 名前 | 雰囲気 | accent |
|---|---|---|
| `claude` | デフォルト。Anthropic の温かいクリーム地（`#faf9f5`）＋クレイ。ダーク表紙・結び（sandwich）＋カードのソフトシャドウ | `#d97757` |
| `minimal` | 洗練ミニマルな白基調・中央バランス | `#c2410c` |
| `midnight` | 紺ヒーロー×ロイヤルブルー。経営・戦略・エグゼクティブ報告向け（sandwich＋シャドウ） | `#3d52c4` |
| `terracotta` | テラコッタ×サンド×セージ。講義・教育・コミュニティの温かさ（sandwich＋シャドウ） | `#b85042` |
| `teal` | 深ティール×ミント。医療・環境・公共の信頼感（sandwich＋シャドウ） | `#00907f` |
| `cherry` | チェリーレッド×ネイビー・セリフ見出し。キーノート・強い主張（sandwich＋シャドウ） | `#c31126` |
| `beamer` | LaTeX beamer(Madrid) 紺。frametitle 帯・定理ブロック・3セルフッターバー | `#b8860b` |
| `tmu-cs` | 白地＋TMU グリーンの学術テーマ。緑見出し＋細下線・左揃えタイトル。[marp-theme-tmu-cs](https://github.com/taishi-n/marp-theme-tmu-cs) (MIT) 移植 | `#006543` |
| `research` | PowerPoint マスタ灰の研究審査テーマ。本文＝罫線＝`#404040`・高密度・上詰め。[marp-theme-dev](https://github.com/katsuzakitomohiro/marp-theme-dev) (MIT) 移植。プリセット「研究審査（高密度）」と相性◎ | `#dd5400` |

### academic 系パレット（配色のみ）

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
marp-pptx convert slides.md            # claude（デフォルト）
marp-pptx convert slides.md -p minimal # 白基調
marp-pptx convert slides.md -p navy    # 紺
```

## 発表者ノート

`<!-- note: ... -->` をスライド内に書くと、その内容が PPTX の**ノート欄**に入る。
複数書くと結合される。本文には表示されない。

```markdown
<!-- _class: result -->
# 実験結果
- 提案手法が全条件で最良
<!-- note: 有意差は p<0.01。質問が来たら付録のablationを見せる -->
```

## Web UI（ブラウザ編集）

`pip install -e ".[web]"` の上で `marp-pptx serve`（既定 `127.0.0.1:8080`、`--host` / `--port`）。

- **プリセットから開始**: 厳選スターターデッキ（最小雛形 / 学術発表 / プロダクト紹介 / 講義・勉強会）をワンクリックで読み込み（`data/presets/`）
- **型ギャラリー**（`/types-page`）: 全52型をサムネ付きで一覧・検索。カードをクリックするとその型でエディタが開く
- **フォーム編集**: 型を選んでフォームに入力 → Markdown を自動生成
- **ライブプレビュー**: 各スライドを PNG で即時レンダ（互換性のため数式は内部的に `png` で表示）
- **MD の保存／読み込み・オートセーブ**、スライドの並べ替え・削除
- **画像アップロード**: ドラッグ＆ドロップで `assets/` に取り込み、PPTX に埋め込み
- **PPTX 読み込み（PPTX → MD）**: 既存 PPTX からテキスト・構造をベストエフォート抽出（学習データ作成用）

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
- **数式** `$$...$$` は display、`$...$` は inline。`--math omml`（既定）は OMML 変換に **pandoc が必要**
  （無い場合は自動で matplotlib PNG にフォールバック）。LibreOffice/Keynote で開くなら最初から `--math png` 推奨
- **日本語フォント**：CSS の `--font-ea` で指定したフォントが自動適用される
- **テンプレート例**は `src/marp_pptx/data/templates/` の 52 ファイルに実例あり

## プログラムから使う（Python API）

```python
from pathlib import Path
from marp_pptx.theme import ThemeConfig, get_default_theme_path, get_palette_path
from marp_pptx.parser import parse_marp
from marp_pptx.builder import PptxBuilder

tc = ThemeConfig.from_css(get_default_theme_path())
tc.apply_palette(get_palette_path("navy"))   # CLI 既定と同じ見た目にするなら "claude"
# tc.math_mode = "png"     # LibreOffice/Keynote 用（既定は "omml"）
# tc.density = "keynote"   # 投影向け

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
