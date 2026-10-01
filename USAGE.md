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

`src/marp_pptx/data/templates/` にある 64 種類の MD ファイル（型ごとに 1 つ）は、
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
| `` `code` `` | モノスペースフォント（行内） | |
| ```` ```json ```` … ```` ``` ```` | コードブロック（淡色パネル・等幅・言語ラベル） | 通常スライド・段組の本文にそのまま書ける。json / geojson / jsonl はキー・文字列・数値を色分け |
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

### PowerPoint 側の土台（スライドマスタ）

出力 PPTX のスライドマスタ・レイアウト・テーマは、デッキのテーマから作り直している。
PowerPoint で開いて編集するときに次のように効く:

| 項目 | 内容 |
|---|---|
| テーマの色 | 配色＝パレット（文字・背景・面・主色・グラフ用の accent1〜6）。色の選択肢にデッキの色が並び、PowerPoint で挿入したグラフや図形もデッキの色になる |
| テーマのフォント | 見出し・本文・日本語フォント＝デッキのフォント。新しいテキストボックスが Calibri / ＭＳ Ｐゴシック にならない |
| レイアウト | 16:9 の 6 種（タイトル スライド / タイトルとコンテンツ / セクション見出し / 2 つのコンテンツ / タイトルのみ / 白紙）。「新しいスライド」で追加したスライドもデッキと同じ位置・書式 |
| スライドタイトル | 各スライドの H1 はタイトルのプレースホルダー。アウトライン表示・ナビゲーター・アクセシビリティチェックでタイトルとして扱われる |
| ページ番号 | スライド番号フィールド。並べ替えても番号が追従する |
| セクション | `divider` ごとに PowerPoint のセクション（スライド一覧が章ごとにまとまる） |
| 文書プロパティ | タイトル＝表紙の H1、形式＝ワイド画面 |

### 日本語フォント

自動対応しています。CSS `--font-ea` で指定したフォントが全ての `<a:ea>` 属性に注入され、
英数字は `--font-body` が適用されます。混在文は同一テキストボックス内で自然に描画されます。

## 基本原則: 「型」を選んで書く

このツールの核心は **64種類のセマンティックなスライド型**。
ユーザーが「何を伝えたいか」を聞いたら、まず **どの型を使うか** を決める。

### 型選択の思考フロー

| ユーザーの意図 | 選ぶべき型 |
|---|---|
| 始まり | `title` |
| 表紙をサムネで判別できるように（写真つき） | `title-figure` |
| 章の区切り | `divider` |
| 予定・目次 | `agenda` |
| 2つを並列比較 | `cols-2` |
| 3つを分類 | `cols-3` |
| 1枚に2〜4トピックを高密度で積む | `sections` |
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
| 定理・定義・例をブロックで並べる（beamer 流） | `blocks` |
| 1つの数式 | `equation` |
| 連立式・最適化問題 | `equations` |
| 図＋キャプション | `figure` |
| 論文図を最大サイズで見せる（図が主役） | `figure-full` |
| 図を左右に寄せ、リード＋解説＋まとめ帯で完結させる | `figure-story` |
| 図＋注釈 | `annotation` |
| 構造図 | `diagram` |
| ブロック図・機器構成図・反復ループ図（mermaid 記法） | `flow` |
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
| 課題→手法→成果を一枚絵で（グラフィカルアブストラクト） | `graphical-abstract` |
| まとめ | `summary` |
| キーメッセージ | `takeaway` |
| 1文を大きく言い切る／論点転換 | `statement` |
| 色面に白抜き主張＋右に本文（キーノート級の1枚） | `split-panel` |
| 1つの数字を主役に | `big-number` |
| 数値データをグラフで（編集可） | `chart` |
| 参考文献 | `references` |
| 紹介論文の書誌・選定理由・要点（輪読/文献調査） | `paper` |
| 関連研究を手法ごとに分類し、各分類の下に論文と課題を並べる | `survey` |
| 論文の一節を原文のまま引用し、読み（解釈）を添える | `excerpt` |
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

#### title-figure — 写真つき表紙

サムネイルで中身が分かる表紙にするとき（写真を下/上/左/右に半面ブリード、full で全面）

```markdown
<!-- _class: title-figure -->
<!-- _side: bottom -->
<!-- source: 地理院タイル（2026-08 撮影） -->

# 公開地理データからの変電所内部構成の実証的機械生成

## Evidence-Paired Substation Single-Line Diagramming

著者名 $^{1}$, 共著者名 $^{2}$

$^{1}$ 所属大学 / 学部 &emsp; $^{2}$ 所属機関 &emsp; 2026年 X月 X日

![](figures/cover.png)

<div class="caption">図: 実出力例 — 公開衛星写真に推定結果を重畳</div>
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
<li>64種類の意味的な型</li>
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

#### graphical-abstract — グラフィカルアブストラクト

表紙直後に課題→手法→成果を1枚の図で示すとき（研究発表の定番）

```markdown
<!-- _class: graphical-abstract -->

# 本研究の全体像

## 一枚で｜問い → 課題 → 提案 → 成果

<div class="ga-top">
  <span class="ga-label">問い</span>
  <span class="ga-body">PV 連系申請の受入可否を、**当日中に**回答できるか？ — 総当たり計算 94 分の壁を、精度を落とさずに破れるかを検証する。</span>
</div>

<div class="ga-problem">
  <span class="ga-label">課題</span>
  ![w:400](figures/duck-curve.png)
  <span class="ga-body">総当たり AC × 二分探索で **94 分**。回答は翌日持ち越し。</span>
</div>

<div class="ga-method">
  <span class="ga-label">提案</span>
  ![w:400](figures/loop.png)
  <span class="ga-steps">感度行列 → LP 一括 → AC 検証</span>
  <span class="ga-body">速さと正しさを ==分業== する。</span>
</div>

<div class="ga-result">
  <span class="ga-label">成果</span>
  <span class="ga-kpi">47×</span>
  <span class="ga-body">94.1 分 → 2.0 分（1,200 ノード実測）
誤差 ±1.8%・平均 3 反復で収束</span>
</div>

<div class="ga-foot">実測条件: Xeon w5-2455X 単スレッド／6.6 kV 放射状フィーダ 300〜1,200 ノード</div>
```

#### figure-full — 論文図の全面表示

論文の特徴的な図を余白0.25inまで最大サイズで見せるとき（図が主役の1枚）

```markdown
<!-- _class: figure-full -->
<!-- source: Vaswani et al. (2017), Fig. 1 -->

# Transformer の全体アーキテクチャ

![w:1200](figures/transformer-architecture.png)

<div class="caption">エンコーダ・デコーダとも自己注意＋FFN の積層のみで構成される</div>
```

#### figure-story — 図の完全解剖

図を左右どちらかに寄せ、リード文（h2）・横の解説・下のまとめ帯で1枚を完結させるとき

```markdown
<!-- _class: figure-story -->
<!-- source: Vaswani et al. (2017), Figure 1 -->

# アーキテクチャが並列化を解放する

## 読み方｜再帰が無いから、全時刻を同時に計算できる

![w:600](figures/architecture.png)

<div class="fs-points">
- **左**: エンコーダ — 自己注意＋FFN を $N=6$ 積層
- **右**: デコーダ — マスク付き自己注意で未来を遮蔽
- 系列方向の依存が無く、**全トークンを同時計算**
- 位置情報は正弦波の位置符号で注入（学習不要）
</div>

<div class="fs-conclusion">まとめ: 逐次計算の除去こそが本質 — 学習 3.5 日（8 GPU）での SOTA 到達は、この構造選択の直接の帰結。</div>
```

#### split-panel — ハーフブリード色面パネル

画面端まで塗った色面に主張を白抜きし、右に本文を置くとき（キーノート級の1枚）

```markdown
<!-- _class: split-panel -->

# 速さと正しさは、もう交換条件ではない

## PROPOSAL

<div class="sp-body">
- 感度行列 LP が **1 秒未満** で候補を出す
- AC 潮流は違反時のみ — 平均 ==3 往復== で収束
- 判定誤差は AC 比 ±1.8%、当日回答が標準になる
- 既存の系統データベースはそのまま使える
</div>
```

#### paper — 文献報告の書誌カード

輪読・文献調査で紹介論文の書誌情報・選定理由・要点を1枚にするとき

```markdown
<!-- _class: paper -->

# Attention Is All You Need

## 文献報告｜再帰を捨てた自己注意のみの系列変換

<div class="pp-meta">
  <span class="pp-authors">Vaswani, A., Shazeer, N., Parmar, N. et al.（Google Brain / Google Research）</span>
  <span class="pp-venue">NeurIPS 2017</span>
  <span class="pp-stats">被引用 130,000+ ／ arXiv:1706.03762</span>
</div>

<div class="pp-why">
  <span class="pp-why-label">選定理由</span>
  <span class="pp-why-body">自研究のスパース注意機構の出発点。計算量 $O(n^2)$ がどの設計判断から生じたかを原典で確認し、削減余地を特定する。</span>
</div>

<div class="pp-points">
- 再帰・畳み込みを排し **Multi-Head Self-Attention** のみで系列変換を構成
- 位置情報は正弦波の位置エンコーディングで注入（学習不要）
- WMT14 EN-DE で BLEU ==28.4==（当時 SOTA）、学習コストは従来比 1/4
</div>
```

#### board — 区画ボード

定義帯・カード・コード・図・番号付き手順・結論帯・出典を、色付きの区画パネルに組んで1枚に詰めるとき。`report` テーマ（`-p report`：紺の見出し＋二色の罫線をスライドマスタのレイアウトに持つ）と組むと研究報告の体裁になる。

| 書き方 | 描画 |
|---|---|
| `<div class="defn">**用語**<br>定義</div>` | 定義帯（淡青の全幅帯・用語は青太字） |
| `<div class="row">` の中に `<div class="col w2">` / `<div class="panel maroon w1">` | 横並びの列（`wN` で幅比、`panel` は左に色帯の区画、色は blue / maroon / navy） |
| 列の中の `#### ラベル` | 列の見出し（直後がコードなら青） |
| 列の中の `### :point: 見出し` ＋ 本文 | カード（`:point:` `:line:` `:polygon:` `:network:` `:arrow:` で図形アイコン） |
| コードフェンス | 暗色のコードパネル（json / geojson はキー・文字列・数値を色分け、列の高さまで伸びる） |
| `- **ラベル** 値` だけの箇条書き | ラベルのチップ＋値の行 |
| 画像の直後の文 | キャプション（小さい灰色） |
| `<div class="callout">→ 文<br>==強調==</div>` | 結論カード（先頭の → で矢印） |
| `<div class="steps">` 内の `### 見出し ｜ 式` ＋ 本文 | 番号付きの段（上辺に色帯・番号丸・右にアイコン） |
| `<div class="band maroon">**要点** — 文<br>**但し書き** — 文</div>` | 強調帯（1 行目の太字は紺、2 行目以降はえんじ） |
| `<!-- _kicker: 章 ｜ 話題 -->` | 見出しの上の小見出し（report テーマ） |

文字は参考にした手作りデッキと同じ大きさ（カード見出し 19pt・本文 15pt・コード 14pt）から始め、入らなければ全体を縮める。空いた高さは行とコード・画像・手順に配る。

```markdown
<!-- _class: board -->
<!-- _kicker: 基礎 ｜ データ形式 -->

# GeoJSON は「地理を JSON で書く」地理フォーマット

<div class="defn">

**GeoJSON（Geographic JSON・RFC 7946）**
JSON の書き方で「点・線・面」の地物と属性を表す。座標は [経度, 緯度] の順で並べる。

</div>

<div class="row">
<div class="col">

#### 3 つのプリミティブ

### :point: Point ｜ 点
変電所・発電所などの地点

### :line: LineString ｜ 線
送電線の経路（頂点の並び）

### :polygon: Polygon ｜ 面
供給エリア・発電所の敷地

<div class="callout">→ 系統の構成要素は すべて 点・線・面 に集約<br>==地理データのまま 系統モデルへ==</div>

</div>
<div class="col">

#### 例：変電所を GeoJSON で書くと

```geojson
{
  "type": "Feature",
  "geometry": {
    "type": "Point",
    "coordinates": [136.22, 36.06]
  },
  "properties": {
    "name": "福井変電所",
    "voltage": 500
  }
}
```

</div>
</div>

<!-- source: 出典：RFC 7946 The GeoJSON Format (IETF, 2016) -->
```

#### survey — 関連研究マップ

関連研究を手法ごとに分類し、該当論文・課題を並べるとき。文献のある分類は 1 行が左右 2 ゾーン（左: 手法名と件数、右: 課題と文献）になり、行の間に罫線が入る。文献の無い分類（手法→性質→課題の整理）は全幅で縦に積む。`<!-- _mark: キーワード -->`（カンマ区切りで複数可）は文献・引用中の語を大小文字を問わずマーカー表示する。`sv-body`（手法の性質）/ `sv-gap`（課題）/ 文献リストはどれも省略可。文献を `著者, "題目," 誌名, 年.` の形で書くと、著者・題目・誌名が濃淡で描き分けられ、引用符は “ ” に、年は前の語と一緒に折り返す。件数は文献のある分類が 2 つ以上のときだけ出る。量が少なければ文字を大きくし、余った高さは行の間に配る。多ければ文献を 9pt まで縮め、それでも入らなければ「手法名の下に文献」を 2 段組で積む。

```markdown
<!-- _class: survey -->
<!-- _mark: distributionally robust -->

# 関連研究｜ネットワーク制約の扱い

## 共通の短所｜緩和・近似モデルではネットワーク制約の充足が保証されない

<div class="sv-group">
<span class="sv-label">電力潮流方程式の**線形近似**モデル</span>
<span class="sv-gap">近似誤差の分だけ電圧・電流の制約を逸脱しうる</span>

- Y. Wen et al., "Centralized distributionally robust chance-constrained dispatch of integrated transmission-distribution systems," IEEE Trans. Power Syst., 2024.
- S. Talari et al., "Sequential clearing of network-aware local energy and flexibility markets in community-based grids," IEEE Trans. Smart Grid, 2024.
- X. Shi et al., "Day-ahead distributionally robust optimization-based scheduling for distribution systems with electric vehicles," IEEE Trans. Smart Grid, 2023.
- Y. Zhou et al., "Aggregated feasible active power region for distributed energy resources with a distributionally robust joint probabilistic guarantee," IEEE Trans. Power Syst., 2025.
</div>

<div class="sv-group">
<span class="sv-label">電力潮流方程式の**二次錐緩和**モデル</span>
<span class="sv-gap">緩和が厳密にならない条件では解が物理的に実現できない</span>

- T. Jiang et al., "Flexibility clearing in joint energy and flexibility markets considering TSO-DSO coordination," IEEE Trans. Smart Grid, 2023.
</div>

<div class="sv-group">
<span class="sv-label">アフィン方策で**緩和**＋機会制約を**近似**</span>

- M. Rayati et al., "Distributionally robust chance constrained optimization for providing flexibility in an active distribution network," IEEE Trans. Smart Grid, 2022.
</div>

<div class="sv-verdict">未解決: 不確実性の下でネットワーク制約の充足を保証しつつ、計算負荷を抑える方法</div>
```

#### excerpt — 原文抜粋＋読み

論文の一節をそのまま `ex-quote` に引き、その下の `ex-read` に解釈を添えるとき。`ex-cite`（省略可）は引用の位置（`Sec. I` など）を引用の後ろに小さく出す。`ex-source` は脚注帯に完全な書誌を出す。`ex-verdict` / `sv-verdict` は `<br>` で改行でき、`**利点**` のように行頭を太字にできる。`check_deck_against_source` / `from-paper` は `ex-quote` が論文本文に逐語で存在するかを照合する（`…` の前後は別々に照合）。

**論文から切り抜く**: スライドに `<!-- _paper: papers/zhou2025.pdf -->`（MD からの相対パス）を書くと、各 `ex-quote` を PDF 内で探し、その段落を**論文の組版のまま画像で切り抜いて**載せる。引用した語にマーカー、前後 1 行は薄く、下に `Sec. I, p. 2` のような位置、右に読み。PDF に見つからない引用だけは文字のカードに戻り、警告が出る（引用の打ち間違いにも気づける）。PyMuPDF（`pip install "marp-pptx[ingest]"`）が必要。

```markdown
<!-- _class: excerpt -->

# 研究課題（補足）｜DER の aggregation

## 他アプローチ｜フレキシビリティを集約して扱う（Zhou et al., 2025）

<div class="ex-item">
<span class="ex-quote">"… directly incorporating a large population of DERs in system-wide scheduling introduces high computing efforts."</span>
<span class="ex-read">DER を個別に直接制御すると==計算負荷が大きい==</span>
</div>

<div class="ex-item">
<span class="ex-quote">"Another promising solution is DER aggregation, which requires the identification of the Aggregated Feasible Power Region (AFPR) … while respecting the network constraints."</span>
<span class="ex-read">集約した実行可能電力領域 AFPR（＝ネットワーク制約を充足する領域）を同定する</span>
</div>

<div class="ex-verdict">**利点**　ネットワーク制約を満たしたまま TSO-DSO 協調に使える<br>**課題**　==不確実性下==での制約充足と disaggregation（計算負荷が大きい）</div>

<div class="ex-source">Y. Zhou, C. Essayeh and T. Morstyn, "Aggregated feasible active power region for distributed energy resources with a distributionally robust joint probabilistic guarantee," IEEE Trans. Power Syst., vol. 40, no. 1, pp. 556–571, 2025.</div>
```

#### sections — 高密度トピック積層

1枚に2〜4トピックを「色付きリード行＋本文」の帯で積むとき（公聴会流の高密度）

```markdown
<!-- _class: sections -->

# 需給運用における課題

<div class="sec">
  <span class="sec-title">ダックカーブ現象｜① 供給側の負担増加</span>
  <span class="sec-body">PVS の大量導入で **正味電力需要** がダックカーブ形状へ変化。正午の谷と夕方の急峻な立ち上がりに、発電機側の調整力だけでは追随できなくなりつつある。</span>
</div>

<div class="sec">
  <span class="sec-title">需給逼迫｜② 予備力の不足</span>
  <span class="sec-body">下げしろ・上げしろの両方向で予備力が逼迫。発電機出力の下限値と最大変化率が制約となり、==需要側の参加== が不可欠になる。</span>
</div>

<div class="sec">
  <span class="sec-title">要請｜③ 意思決定の定量的根拠</span>
  <span class="sec-body">技術選択・設備容量計画・運転スケジューリングの 3 つの意思決定に、評価指標（コスト・CO₂ 排出量）の定量値を与える枠組みが必要。</span>
</div>
```

#### blocks — 定理・定義ブロック

beamer 流の定理環境（theorem/example/alert）を並べるとき

```markdown
<!-- _class: blocks -->

# 定理ブロック（beamer 流）

<div class="bk-container">

<div class="bk theorem">
  <span class="bk-title">定理 1（収束性）</span>
  <span class="bk-body">提案する反復法は任意の初期点から線形レートで収束する。すなわち $\|x_{k+1} - x^*\| \le \rho \|x_k - x^*\|$（$0 < \rho < 1$）。</span>
</div>

<div class="bk example">
  <span class="bk-title">例（二次関数）</span>
  <span class="bk-body">$f(x) = x^2$ では $\rho = 1/2$ となり、10 反復で誤差はおよそ $10^{-3}$ 倍に縮む。</span>
</div>

<div class="bk alert">
  <span class="bk-title">注意</span>
  <span class="bk-body">ステップ幅が $2/L$ を超えると発散する（$L$ は勾配のリプシッツ定数）。</span>
</div>

</div>
```

#### flow — ブロック図・ループ図

mermaid flowchart 記法から編集可能なブロック図/機器構成図/反復ループ図を描くとき

```markdown
<!-- _class: flow -->

# 提案手法の全体構成

## 反復補正｜LP が候補を出し、AC 潮流が正しさを保証する

```mermaid
flowchart LR
  S[感度行列 S<br>ヤコビアンから構築] --> LP[線形最適化 LP<br>全ノード一括]:::accent
  LP -->|候補解 p*| AC[AC 潮流計算<br>電圧・電流を検証]
  AC -->|違反なし| OUT([受入可否を回答]):::primary
  AC -.->|違反あり・再線形化| LP
```
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
- **型ギャラリー**（`/types-page`）: 全64型をサムネ付きで一覧・検索。カードをクリックするとその型でエディタが開く
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
