---
marp: true
---

<!-- _class: title -->

# 2 つの Marp テーマを移植し、7 つのバグを直した
## marp-pptx 進化レポート

Claude Code × 重信 — marp-pptx プロジェクト

2026-06-12 / セッションレポート

<!-- note: 本デッキ自体が新テーマ tmu-cs のデモ。タイトルの左揃えは今日実装した title_align: left がそのまま出ている。 -->

---

<!-- _class: agenda -->

# 本日の内容

<div class="agenda-list">

1. 今日増えたもの — テーマ 2・プリセット 1
2. 移植の手順 — 2 ファイルで完結する
3. 直したバグ 7 件
4. 数字で見る品質
5. 残る課題と次の一手

</div>

---

<!-- _class: statement -->

外部 Marp テーマの移植は、CSS と YAML の 2 ファイルで完結する。

---

<!-- _class: cols-2 -->

# 新テーマ｜tmu-cs と research

<div class="columns">
<div>

## tmu-cs — 都立大グリーン

- 出典: taishi-n/marp-theme-tmu-cs（MIT）
- 白地に TMU 緑 `#006543`、見出しはすべて緑
- 緑見出し＋細下線、**左揃えタイトル・上詰め**
- 結論ボックスは note 地＋緑の左線（blockquote 流儀）

</div>
<div>

## research — PowerPoint マスタ灰

- 出典: katsuzakitomohiro/marp-theme-dev（MIT）
- 本文＝罫線＝`#404040`、太字も本文色のまま
- 茶 `#dd5400` は**キー数値だけ**に使う設計契約
- 高密度・コンパクト。研究審査・予備審査向け

</div>
</div>

<div class="footnote">このスライドの左右カラムの上端が揃っているのは、今日直したバグ #4 の効果。昨日まではズレていた。</div>

---

<!-- _class: zone-process -->

# 移植手順｜4 ステップで -p が通る

<div class="zp-container">

<div class="zp-step">
  <span class="zp-num">1</span>
  <span class="zp-title">色の対応付け</span>
  <span class="zp-body">大元 CSS の変数を palettes/<name>.css の --color-* に写す。出典とライセンスを冒頭に明記。</span>
</div>

<div class="zp-step">
  <span class="zp-num">2</span>
  <span class="zp-title">レイアウト文法</span>
  <span class="zp-body">config-<name>.yaml に h1_deco / title_align / vertical_align / box_style を書く。色だけでなく組版の癖ごと移す。</span>
</div>

<div class="zp-step">
  <span class="zp-num">3</span>
  <span class="zp-title">露出</span>
  <span class="zp-body">CLI themes に 1 行追加。Web UI は palettes/ を自動列挙するので作業不要。</span>
</div>

<div class="zp-step">
  <span class="zp-num">4</span>
  <span class="zp-title">検証</span>
  <span class="zp-body">devrender.sh でレンダ → visual_lint＋collision 検出＋目視。既存テーマはピクセル比較で不変を確認。</span>
</div>

</div>

---

<!-- _class: before-after -->

# Web UI のパレット欄が「全テーマ」になった

<div class="ba-before">
  <span class="ba-label">Before</span>
  <span class="ba-body">academic 系 10 色だけを列挙。claude / minimal すら選択肢に出ず、テーマを増やしても誰も気づけない。既定の表記も「Default (mono)」のまま古かった。</span>
</div>

<div class="ba-after">
  <span class="ba-label">After</span>
  <span class="ba-body">palettes/ の全 CSS を自動列挙。tmu-cs / research を含む 14 テーマから選べる。新しいテーマを置くだけで一覧に出る。</span>
</div>

---

<!-- _class: table-slide -->

# 直したバグ｜7 件すべてテストで固定

| # | 症状 | 修正 |
|---|------|------|
| 1–2 | diagram / chart の長キャプションが折り返されず画面外へ | word_wrap ＋ 折返し見積りで高さを動的化 |
| 3 | `title_align` が実装されていない死にトークン | left を実装（claude はピクセル差 0 で不変） |
| 4 | cols-2 の左右カラムの上端がズレる | 全カラム共通 top ＋ 折返し考慮の見積り |
| 5 | PNG モードで `\mathrm{pv}` が「mathrmpv」に劣化 | コマンド剥がしを strip の前に追加 |
| 6 | references が上詰めのまま（中央化の漏れ） | 縦中央化 ＋ lint の偽陽性も解消 |
| 7 | Web UI が academic 系しか列挙しない | 全テーマ自動列挙に修正 |

<div class="footnote">全件 tests/test_regressions.py にリグレッションテストとして固定。#4 は型ギャラリーのカタログ画像にも写っていた古参バグだった。</div>

---

<!-- _class: chart -->
<!-- _chart: column -->

# テストは 4 日で 94 → 125 に増えた

| 日付 | テスト数 |
|---|---|
| 6/8 | 94 |
| 6/9 | 116 |
| 6/10 | 119 |
| 6/12 | 125 |

<div class="chart-caption">実コンテンツ検証 → 自己改善ループ → テーマ移植。バグを直すたびにリグレッションテストを積み増した結果。</div>

---

<!-- _class: kpi -->

# 数字で見る今日のセッション

<div class="kpi-container">

<div class="kpi-item">
  <span class="kpi-value">2</span>
  <span class="kpi-label">移植テーマ（tmu-cs / research）</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">16 枚</span>
  <span class="kpi-label">新プリセット「研究審査（高密度）」</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">7 件</span>
  <span class="kpi-label">発見・修正したバグ</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">125</span>
  <span class="kpi-label">テスト全通過</span>
</div>

</div>

---

<!-- _class: pros-cons -->

# 棚卸し｜できたこと・残っていること

<div class="pc-pros">
<li>外部テーマ移植の「型」が確立 — 次は 2 ファイル＋1 行で増やせる</li>
<li>色だけでなくレイアウト文法ごと移植（左揃え・上詰め・抑制強調）</li>
<li>beamer 風高密度プリセットが Web UI から 1 クリックで使える</li>
</div>

<div class="pc-cons">
<li>カラム内の box / box-accent が箱として描画されない（skeleton と乖離）</li>
<li>figure-cols / gallery-img のキャプション wrap 漏れが残る</li>
<li>Noto Sans JP / Segoe UI 不在の環境はヒラギノ等へフォールバック</li>
</div>

---

<!-- _class: takeaway -->

# キーメッセージ

<div class="ta-main">テーマは 2 ファイルで増やせる。質は 125 のテストが守る</div>

<div class="ta-points">
<li>tmu-cs / research は配色だけでなく大元の設計契約ごと移植した</li>
<li>移植の途中で踏んだバグ 7 件は、その場で直してテストに固定した</li>
<li>次の一手は box 描画と caption wrap の残り 2 箇所</li>
</div>

---

<!-- _class: dark -->

テーマは設定で増える。品質はテストで増える。

---

<!-- _class: end -->

# ありがとうございました

このデッキの再現: marp-pptx convert demo/theme-port-report.md -p tmu-cs

marp-pptx v0.3.1 / feature/web-ui
