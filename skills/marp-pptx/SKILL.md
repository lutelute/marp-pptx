---
name: marp-pptx
description: "marp-pptx で「編集可能な PowerPoint(.pptx)」を作るためのオーサリングガイド。以下の場面で使用: (1) スライド・プレゼン・デッキ・発表資料を marp-pptx で作るとき、(2) Markdown から編集可能な PPTX を生成するとき、(3) 52種のセマンティック型(kpi/equation/timeline/sandwich/statement 等)からスライドを構成するとき、(4) 学術発表・講義・プロダクト紹介などの資料を素早く組むとき、(5) 既存 Marp Markdown を marp-pptx 形式に直すとき。一枚絵ではなく実テキストボックス＋実テーブル＋OMML数式で出力する。"
---

# marp-pptx オーサリングガイド

**marp-pptx** は Marp Markdown を **完全に編集可能な PowerPoint** に変換するツール。
画像化された一枚絵ではなく、実テキストボックス・実テーブル・OMML 数式（PowerPoint で編集可）を出力する。
価値の核は **52 種のセマンティックなスライド型**。「何を伝えたいか」から型を選んで書く。

## ワークフロー

1. **目的を確認** — 何を・誰に・どう伝えるか。スライド枚数の目安
2. **型で構成設計** — 各スライドに型を割り当てる（下の「型選択フロー」）
   - 典型: `title → agenda → rq → sandwich → kpi/result → takeaway → end`
3. **Markdown を書く** — フロントマター + 型ごとの HTML 構造（下記＆ `references/type-skeletons.md`）
   - 文字量は「1行あたりの実測上限」に収める（下の「文字数の目安」）
4. **変換** — `marp-pptx convert deck.md -o deck.pptx`
5. **検証（必須）** — `marp-pptx doctor deck.md`。**error は必ず直してから出す**
6. **確認** — スライド数・画像・数式が想定通りか。仕上げは PowerPoint で

## 検証（doctor）— 出す前に必ず

```bash
marp-pptx doctor deck.md          # .md ならビルドしてから検査／.pptx も可
marp-pptx doctor deck.md --json   # 機械可読
marp-pptx doctor deck.md --strict # warn 以上で exit 1（CI 用）
```

各テキストボックスを**実フォントの字送り**で測り、レンダラ無し・API キー無しで
「読み手が気づく欠陥」だけを返す。LibreOffice も不要で数十 ms。

| kind | 意味 | 対応 |
|---|---|---|
| `overflow` | 文字がボックスに入っていない（**最頻の欠陥**） | 文を削る／スライドを分割。提示された pt まで下げるのは最後の手段 |
| `overlap` | テキスト同士が重なる／近すぎる | 上のブロックを短くする |
| `offslide` | スライド外／端に寄りすぎ | 内容を減らす（座標は型が決める） |
| `contrast` | WCAG AA 未満の配色 | 濃い色に変える／パレットを替える |
| `font` | 未インストール／プレビューが当てにならない書体 | 安全な書体（Arial/Calibri/Cambria/Times New Roman）へ |
| `package` | PowerPoint が「修復が必要」と言う破損 | 画像形式・パスを確認して再ビルド |
| `deck` | 同じレイアウトの連続・図の無いスライド | 型を変えて緩急をつける |

**severity=error はそのまま出さない。** `overflow` は文字を減らして直すのが第一手
（フォントを縮めるのは読みにくさに直結する）。

MCP 接続時は **`check_deck(markdown=...)`** が同じ検査を返す。
`preview_png` より先にこれを回す（速く・正確で・環境依存が無い）。

## 文字数の目安（claude テーマ / font_scale=1.0 の実測）

| 場所 | 1 行に入る量（日本語 / 英字） | 備考 |
|---|---|---|
| H1（スライド見出し, 34pt） | **25 字 / 52 字** | 見出し帯は 1 行分。超えると自動縮小（〜30pt で 29 字）→2 行化 |
| 本文・箇条書き（18pt, 全幅） | 48 字 / 101 字 | 本文領域は最大 16 行 |
| 2 カラム内（16pt） | 26 字 / 55 字 | `cols-2` / `sandwich` の各カラム |
| カード本文（15pt, 3 列） | 15 字 / 33 字 | `zone-*` / `card-grid`。**ここが一番狭い** |
| カード見出し（18pt, 3 列） | 13 字 / 27 字 | 名詞句で。文にしない |
| キャプション（14pt, 全幅） | 62 字 / 129 字 | 図表の下 |
| KPI 値（44pt, 3 列） | 6 字 / 13 字 | `97%` `10x` のような短い値 |

（claude テーマ・font_scale=1.0・Hiragino Sans / Helvetica Neue 実測値。
英字は一般的な文章の平均字幅で換算。）

超えたぶんは折り返す（悪くはない）が、**H1・カード見出し・KPI 値は 1 行が設計**。
3 列カードに文章を入れると必ず溢れる — 名詞句に削る。迷ったら doctor に測らせる。

> 速く始めるなら **プリセット**から: `marp-pptx serve` の Web UI 右「プリセットから開始」、
> または `src/marp_pptx/data/presets/*.md`（academic-talk / academic-dense / paper-review / product / lecture / minimal）を雛形に。
> **文献報告（輪読・サーベイ）は `paper-review` プリセット**が流れ一式: 書誌カード(paper)→主張(sections)→
> 手法図(flow)→数式→比較表(表の `◎○△×` は自動で緑/琥珀/赤に着色)→強み限界(pros-cons)→自研究への示唆(takeaway)→文献。

## 論文・リポジトリからデッキを作る（グラウンディング必須）

「この論文/リポを渡すのでスライドにして」と言われたら、**記憶で書かず、必ずソースから起こす**。
MCP 接続時（`marp-pptx[mcp]`）は以下のツールを使う:

1. **`read_paper(pdf_or_arxiv_url)`** — 論文を構造化抽出（title / abstract / sections / figures / **numbers**）。
   - 各スライドは抽出された `sections` の文章から書く。数字（BLEU・精度・速度など）は**必ず `numbers` の値を使い、自分で発明しない**。
   - 図は `figures[].path`（抽出済みPNG）を `![w:760](path)` で貼る。
2. **`read_repo(path)`** — リポを要約（readme / tree / languages / key_files）。「何をするコードか」はここから書く。
3. **下書き** — 抽出内容を `title → agenda → rq → before-after/method → equation → kpi/result → takeaway → references → end` に割り付け（型は下表）。
4. **`check_deck_against_source(markdown, paper_path=...)`** — ★**出す前に必ず実行**。`unsupported` に出た数値は**ソースに無い＝ハルシネ**。論文を見直して直すか、その数値を落とす。`score` が 100 になるまで詰める。
5. **`build_pptx(markdown, ...)`** → **`preview_png(markdown)`** で各スライドを画像確認 → レイアウト崩れ・溢れを直す。

> MCP 無し（CLI のみ）の場合: 論文本文を自分で読み、`marp-pptx convert` で生成。**数値はソース文中に実在するか必ず照合**してから載せる（同じ規律）。
> 図抽出・本文抽出は `pip install "marp-pptx[ingest]"`（PyMuPDF）が必要。
> 人間がエージェント無しで一発生成するなら **`marp-pptx from-paper paper.pdf --repo . -o deck.pptx`**（`marp-pptx[ai]` + `ANTHROPIC_API_KEY`）。上記の取り込み→下書き→数値自動修正→ビルドを内蔵。

## 基本ルール（必須）

```markdown
---
marp: true
---

<!-- _class: title -->
# タイトル
## サブタイトル（キッカー）
発表者 / 2026
```

- スライド区切りは `---`（前後に空行）。**フロントマターは先頭に1つだけ**
- 型指定は各スライド冒頭の `<!-- _class: 型名 -->`。**未指定だと素の箇条書き**になる
- **テーマ/配色は CLI の `-p` で選ぶ**。frontmatter の `theme:` / `math:` は読まれない
- 各型は**特定の HTML 構造**を期待する。`<div class="...">` の入れ子を崩すとその箇所が空になる
- 発表者ノートは `<!-- note: ... -->`（PPTX のノート欄へ）
- **段階開示**: `sections` / `flow` のスライドに `<!-- build -->` を書くと、帯／ノードが
  1つずつ増える連作スライドに自動展開（未来分はゴースト表示）。**導出・学習過程を見せる要**
- `**太字**` / `` `code` `` / `$x^2$`（インライン数式）/ `$$...$$`（ディスプレイ数式）/
  `==マーカー==`（蛍光ペン強調・キーワードを1枚に1〜2個）対応。*斜体* は無効
- 改行のみ＝同一段落（ソフトラップ）、空行＝段落分け

## 型選択フロー（意図 → 型）

| 伝えたいこと | 型 |
|---|---|
| 表紙 / 章区切り / 終わり | `title` / `divider` / `end` |
| 目次・構成 | `agenda` |
| 研究課題・問い | `rq` |
| 2つ比較 / 3つ分類 | `cols-2` / `cols-3` |
| 概要→詳細→結論 | `sandwich`（3条件なら `sandwich-3col`） |
| 賛否・長短 / 2軸評価 | `pros-cons` / `zone-matrix` |
| 時系列（横/縦）/ 手順 / 前後比較 | `timeline-h` / `timeline-v` / `steps` / `before-after` |
| 絞り込み / 積層 / 全体像 / 1つ強調 | `funnel` / `stack` / `overview` / `highlight` |
| KPI・数値 / 単一指標 / グラフ | `kpi` / `big-number` / `chart` |
| 結果（単/二/複） | `result` / `result-dual` / `multi-result` |
| 定義 / 数式 / 連立式 / 図 / 注釈図 / コード | `definition` / `equation` / `equations` / `diagram` / `annotation` / `code` |
| フロー(A→B→C) / 詳細フロー | `zone-flow` / `zone-process` |
| **ブロック図・機器構成図・反復ループ図** | `flow`（```mermaid flowchart 記法→編集可能な図形＋矢印。戻り辺 `-.->` でループ） |
| 引用 / 沿革 / 人物 / 1文断言 | `quote` / `history` / `profile` / `statement` |
| まとめ / キーメッセージ / 文献 / 表 | `summary` / `takeaway` / `references` / `table-slide` |
| **1枚に2〜4トピックを高密度で**（公聴会流） | `sections`（色付きリード行＋本文の帯 ×N） |
| **文献報告・輪読の書誌カード** | `paper`（著者・会議バッジ・被引用・選定理由・要点） |
| **キーノート級の1枚**（色面に白抜き主張＋本文） | `split-panel`（左40%フルブリード色面。h2がキッカー） |
| **グラフィカルアブストラクト**（課題▶提案▶成果の一枚絵） | `graphical-abstract`（3パネル＋太矢印＋大数字＋実測条件フット） |
| **論文の特徴図を最大サイズで**（図が主役） | `figure-full`（余白0.25inまで画像。図未配置でも枠＝プレースホルダが立つ） |
| **図の完全解剖**（リード＋横解説＋まとめ帯） | `figure-story`（図は左右どちらか=`<!-- _side: right -->`で反転。アスペクト比で幅自動配分） |

全52型の意味は `marp-pptx types`（`--json` で機械可読、`-c <category>` で絞り込み）。
**各型の正確な HTML 骨組みは `references/type-skeletons.md` を参照**（崩さず埋める）。

### バリエーション（カタログ外・応用）

`marp-pptx types` には出ないが、ビルダーが解釈する変種:

| `_class` | 用途 | 構造 |
|---|---|---|
| `cols-2-wide-l` / `cols-2-wide-r` | 左/右を広くした 2 カラム（62:38） | `cols-2` と同じ |
| `dark` | `statement` の暗背景版（1文を黒地に） | `statement` と同じ（本文1文） |
| `big-statement` | より大きい `statement` | `statement` と同じ |
| `big-number-dark` | `big-number` の暗背景版 | `big-number` と同じ |

暗背景は各スライド先頭の `<!-- bg: dark -->` ディレクティブでも指定できる。

## よく使う骨組み（抜粋）

```markdown
<!-- _class: agenda -->
# 本日の内容
<div class="agenda-list">
1. 背景
2. 手法
3. 結果
</div>

<!-- _class: kpi -->
# 成果
<div class="kpi-container">
<div><span class="kpi-value">97%</span><span class="kpi-label">精度</span></div>
<div><span class="kpi-value">10x</span><span class="kpi-label">高速化</span></div>
</div>

<!-- _class: equation -->
# ベイズの定理
<div class="eq-main">
$$P(A|B) = \frac{P(B|A)P(A)}{P(B)}$$
</div>
<div class="eq-desc">
<span>$P(A|B)$</span><span>事後確率</span>
<span>$P(B|A)$</span><span>尤度</span>
</div>

<!-- _class: statement -->

本研究は、スパース注意機構に初めて理論的保証を与える。

<!-- _class: big-number -->
<!-- source: 社内ベンチ 2026 -->
# 主要成果
<div class="big-number">
  <span class="bn-value">89.4%</span>
  <span class="bn-label">分類精度</span>
  <span class="bn-caption">従来比 +4.2pt</span>
</div>

<!-- _class: chart -->
<!-- _chart: column -->
# 計算時間
| 系列長 | A | Ours |
|---|---|---|
| 1024 | 1.0 | 0.6 |
| 4096 | 4.2 | 1.1 |
```

## デザイン原則（必読 — 型選択と同じ重み）

スライドの質は「型の選び方」と「パレットの選び方」でほぼ決まる。

**パレットは題材で選ぶ。** 別の題材に流用しても成立する配色は、選び方が浅い。

| 題材・トーン | `-p` |
|---|---|
| 既定・研究/技術全般（cream + clay） | `claude`（無指定） |
| 経営・戦略・エグゼクティブ報告 | `midnight` |
| 医療・環境・インフラ・公共（信頼感） | `teal` |
| 講義・教育・コミュニティ（温かさ） | `terracotta` |
| キーノート・ブランド発表・強い主張 | `cherry` |
| 学会発表・審査（高密度・保守的） | `beamer` / `research` / `tmu-cs` |
| 博士公聴会・審査会（パンくず＋n／m＋図N自動採番＋6色チャート） | `research` |
| 白基調ミニマル | `minimal`（他 `navy`/`forest`/`wine` 等 academic 系） |

claude / midnight / teal / terracotta / cherry は **sandwich 構造**（ダーク表紙→明るい本文→
ダーク結び）とカードのソフトシャドウが自動で入る。暗い1枚を挟みたければ
`statement` + `<!-- bg: dark -->`。

**1スライド1メッセージ + 視覚アンカー1つ。**
- 素の箇条書きだけのスライドを連続させない。比較なら `cols-2`/`before-after`、
  数値なら `kpi`/`big-number`/`chart`、構造なら `zone-flow`/`diagram` に置き換える
- 同じ型を3枚以上連続させない（doctor が `deck` 警告で検出する）
- 誇りたい数字は本文に埋めず `big-number`（1つ）か `kpi`（2〜4つ）に昇格させる
- H1 は「ラベル」でなく**要約**にする（「実験結果」でなく「精度を落とさず 47 倍速い」）。
  短い H1 は大きく描画され、長いと自動縮小で迫力が死ぬ — 結論を短く言い切る

**説得の構造（1枚ごとに3点セット）**
- **主張**（H1）→ **根拠**（数値・図・表。`<!-- source: -->` で出典）→ **解釈**（h2 リード or 結論箱）。
  根拠のない主張スライドを作らない。数値は必ずソース照合（from-paper の fidelity と同じ規律）
- **導出過程を見せる**: 結論をいきなり出さず、`sections`＋`<!-- build -->` で
  観察→仮説→検証の順に積み上げる。仕組みは `flow`＋`<!-- build -->` でノードを1つずつ
- **限界を明示する**: 「言えること／まだ言えないこと」の1枚が主張の信頼を作る
- **表紙の直後に `graphical-abstract` を1枚**: 聴衆は最初の2枚で全体像を得る（研究発表の定番）
- 参照実装: `demo/claude-showcase.md`（導出build・出典付き実測・◎○△×比較・限界明示の20枚）

## 変換コマンド

```bash
marp-pptx convert deck.md -o deck.pptx        # 既定 claude テーマ（cream + clay）
marp-pptx convert deck.md -p midnight         # 題材で選ぶ（上の表）
marp-pptx convert deck.md --math png          # LibreOffice・Keynote で開くなら数式を画像化
marp-pptx convert deck.md --density keynote   # 投影向けに大きめ
marp-pptx convert deck.md --density dense     # 公聴会級の高密度（文字数上限も緩む）
```

- 数式は既定 `--math omml`（PowerPoint で編集可、要 pandoc）。pandoc 不在時や LibreOffice/Keynote 用途は `--math png`
- 画像は MD からの相対パス。`![w:800](assets/x.png)` で幅指定
- 未インストールなら `pip install marp-pptx`（Web UI は `marp-pptx[web]`）

## 落とし穴

- 型を指定し忘れる → 素の箇条書きになる（`<!-- _class: -->` を必ず付ける）
- HTML 構造を崩す → その箇所が空になる（`references/type-skeletons.md` 通りに）
- frontmatter を各スライドに書く → 2枚目以降は frontmatter にならない（先頭1つだけ）
- `theme:` を frontmatter に書いても効かない → 配色は `-p` で
- スライドが想定数と違う → `---` 区切りの前後に空行があるか確認
