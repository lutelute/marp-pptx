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
4. **変換** — `marp-pptx convert deck.md -o deck.pptx`
5. **確認** — スライド数・画像・数式が想定通りか。仕上げは PowerPoint で

> 速く始めるなら **プリセット**から: `marp-pptx serve` の Web UI 右「プリセットから開始」、
> または `src/marp_pptx/data/presets/*.md`（academic-talk / product / lecture / minimal）を雛形に。

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
- `**太字**` / `` `code` `` / `$x^2$`（インライン数式）/ `$$...$$`（ディスプレイ数式）対応。*斜体* は無効
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
| 引用 / 沿革 / 人物 / 1文断言 | `quote` / `history` / `profile` / `statement` |
| まとめ / キーメッセージ / 文献 / 表 | `summary` / `takeaway` / `references` / `table-slide` |

全52型の意味は `marp-pptx types`（`--json` で機械可読、`-c <category>` で絞り込み）。
**各型の正確な HTML 骨組みは `references/type-skeletons.md` を参照**（崩さず埋める）。

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

## 変換コマンド

```bash
marp-pptx convert deck.md -o deck.pptx        # 既定 claude テーマ（cream + clay）
marp-pptx convert deck.md -p minimal          # 白基調 / -p navy 等のパレット
marp-pptx convert deck.md --math png          # LibreOffice・Keynote で開くなら数式を画像化
marp-pptx convert deck.md --density keynote   # 投影向けに大きめ
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
