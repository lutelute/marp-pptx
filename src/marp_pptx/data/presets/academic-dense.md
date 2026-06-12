---
marp: true
---

<!-- _class: title -->

# 太陽光発電・蓄電池を備えた需要家の最適充放電スケジューリング
## 線形計画によるエネルギーマネジメント

山田 太郎 — ○○大学大学院 工学研究科

研究発表会（予備審査） / 2026年◯月

<!-- note: 研究審査向け高密度テンプレ。タイトルは「対象＋手法」を具体的に。発表会名と日付を必ず入れる。 -->

---

<!-- _class: agenda -->

# 発表の流れ

<div class="agenda-list">

1. 背景 — 再エネ大量導入と需要家の役割
2. 目的 — 何を最適化するか
3. モデル化 — 対象システムと変数定義
4. 定式化 — 線形計画問題への帰着
5. 数値実験 — 条件・結果・定量評価
6. 考察とまとめ

</div>

---

<!-- _class: sandwich -->

# 背景｜電気料金の高騰と自家消費ニーズ

<div class="top">

<p class="lead">太陽光発電（PV）と蓄電池（BESS）を持つ需要家にとって、「いつ充電し、いつ放電するか」が電気料金を左右する。運用は依然として単純なルールベースが主流である。</p>

</div>

<div class="columns c2">
<div>

## 現状の課題

- ルールベース運用は料金体系の変化に追従できない
- 余剰 PV の売電と自家消費の最適バランスが不明
- 蓄電池の劣化を考慮しない過剰充放電

</div>
<div>

## 求められる要件

- 翌日 24 時間のスケジュールを**数秒で算出**
- 買電・売電価格、PV・需要予測を反映
- 既存の HEMS/BEMS に組み込める計算量

</div>
</div>

<div class="bottom">

<div class="conclusion">

**本研究の立場**: 充放電スケジューリングを線形計画（LP）に帰着させ、厳密最適解を実用時間で得る。

</div>

</div>

---

<!-- _class: rq -->

# 目的

<div class="rq-main">
PV・蓄電池・系統購入を組み合わせた需要家の電力運用で、売買電料金を最小化するスケジュールを求める
</div>

<div class="rq-sub">
— 予測値（PV 出力・需要・価格）を所与とした day-ahead 最適化として定式化する
</div>

---

<!-- _class: cols-2 -->

# 対象システム｜需要側エネルギーシステムの構成

<div class="columns">
<div>

## 構成要素

- **PV**: 定格 5 kW、売電可
- **BESS**: 容量 10 kWh、充放電効率 95%
- **負荷**: 一般住宅需要（30 分値）
- **系統**: 買電・売電（時間帯別価格）

</div>
<div>

## 前提条件

**外生入力（予測値）**: PV 出力 $P_{\mathrm{pv}}(t)$、需要 $P_{\mathrm{load}}(t)$、買電価格 $c_{\mathrm{buy}}(t)$、売電価格 $c_{\mathrm{sell}}(t)$

**操作量**: 蓄電池の充電電力 $P_{\mathrm{ch}}(t)$ と放電電力 $P_{\mathrm{dis}}(t)$（30 分刻み・48 ステップ）

</div>
</div>

---

<!-- _class: definition -->

# モデル化（1／3）｜SOC — 蓄電池の状態変数

<div class="df-term">充電状態 SOC (State of Charge)</div>

<div class="df-body">蓄電池に蓄えられた電力量を定格容量で正規化した値。時刻 $t$ の SOC は直前の SOC と充放電電力から逐次的に決まり、$\mathrm{SOC}_{\min} \le \mathrm{SOC}(t) \le \mathrm{SOC}_{\max}$ の運用範囲に制約される。</div>

<div class="df-note">関連: 充放電効率 $\eta$、C レート制約、カレンダー劣化・サイクル劣化</div>

---

<!-- _class: equations -->

# モデル化（2／3）｜システムの動的モデル

<div class="eq-system">
  <div class="row">
    <span class="label">SOC 遷移</span>

$$E(t+1) = E(t) + \left(\eta\, P_{\mathrm{ch}}(t) - \tfrac{1}{\eta} P_{\mathrm{dis}}(t)\right) \Delta t$$

  </div>
  <div class="row">
    <span class="label">電力収支</span>

$$P_{\mathrm{grid}}(t) = P_{\mathrm{load}}(t) - P_{\mathrm{pv}}(t) + P_{\mathrm{ch}}(t) - P_{\mathrm{dis}}(t)$$

  </div>
  <div class="row">
    <span class="label">売買分解</span>

$$P_{\mathrm{grid}}(t) = P_{\mathrm{buy}}(t) - P_{\mathrm{sell}}(t), \quad P_{\mathrm{buy}}, P_{\mathrm{sell}} \ge 0$$

  </div>
</div>

<div class="eq-desc">
  <span class="sym">$E(t)$</span>
  <span>蓄電量 [kWh]（$E = \mathrm{SOC} \times E_{\mathrm{rated}}$）</span>
  <span class="sym">$P_{\mathrm{grid}}(t)$</span>
  <span>正味系統電力（正: 買電 / 負: 売電）</span>
</div>

<div class="footnote">正味電力を非負の買電・売電に分解することで、価格の異なる双方向取引を線形のまま扱える。</div>

---

<!-- _class: equations -->

# 定式化（3／3）｜線形計画問題

<div class="eq-system">
  <div class="row">
    <span class="label">minimize</span>

$$\sum_{t=1}^{48} \left( c_{\mathrm{buy}}(t)\, P_{\mathrm{buy}}(t) - c_{\mathrm{sell}}(t)\, P_{\mathrm{sell}}(t) \right) \Delta t$$

  </div>
  <div class="row">
    <span class="label">subject to</span>

$$E_{\min} \le E(t) \le E_{\max}, \quad 0 \le P_{\mathrm{ch}}(t), P_{\mathrm{dis}}(t) \le P_{\mathrm{rate}}$$

  </div>
  <div class="row">
    <span class="label"></span>

$$E(48) = E(0) \quad \text{（周期境界: 翌日へ持ち越さない）}$$

  </div>
</div>

<div class="footnote">決定変数 192 個・制約約 300 本の LP。汎用ソルバで 1 秒未満に厳密最適解が得られる。</div>

<!-- note: 目的関数・制約とも全て線形。整数変数を使わないことが高速性の鍵である点を口頭で強調。 -->

---

<!-- _class: zone-process -->

# 解法｜day-ahead スケジューリングの手順

<div class="zp-container">

<div class="zp-step">
  <span class="zp-num">1</span>
  <span class="zp-title">予測値の取得</span>
  <span class="zp-body">翌日の PV 出力・需要・価格を予測モデルから 30 分値で取得。</span>
</div>

<div class="zp-step">
  <span class="zp-num">2</span>
  <span class="zp-title">LP の構築</span>
  <span class="zp-body">SOC 遷移・電力収支・運用範囲を制約行列に展開。</span>
</div>

<div class="zp-step">
  <span class="zp-num">3</span>
  <span class="zp-title">求解</span>
  <span class="zp-body">単体法／内点法で厳密最適解を取得（1 秒未満）。</span>
</div>

<div class="zp-step">
  <span class="zp-num">4</span>
  <span class="zp-title">スケジュール配信</span>
  <span class="zp-body">充放電計画を HEMS へ送信。当日は実測との偏差を監視。</span>
</div>

</div>

---

<!-- _class: table-slide -->

# 数値実験（1／3）｜実験条件

| 項目 | 設定値 |
|------|--------|
| 対象期間 | 夏季 31 日間（30 分刻み・48 ステップ/日） |
| PV / BESS | 定格 5 kW / 容量 10 kWh（効率 95%） |
| 買電価格 | 昼 35 円/kWh・夜 21 円/kWh（時間帯別） |
| 売電価格 | 8.5 円/kWh（固定） |
| 比較ケース | ①BESS なし ②ルールベース ③提案（LP） |

<div class="footnote">需要・PV 出力は公開データセットの実測プロファイルを使用。予測誤差ゼロ（完全予見）を仮定し、手法の上限性能を評価する。</div>

---

<!-- _class: chart -->
<!-- _chart: column -->

# 数値実験（2／3）｜日別電気料金の比較

| ケース | 平均 [円/日] |
|---|---|
| BESS なし | 412 |
| ルールベース | 358 |
| 提案（LP） | 296 |

<div class="chart-caption">夏季 31 日間の平均。ルールベース比で 17%、BESS なし比で 28% の料金削減。</div>

---

<!-- _class: kpi -->

# 数値実験（3／3）｜定量評価

<div class="kpi-container">

<div class="kpi-item">
  <span class="kpi-value">−28%</span>
  <span class="kpi-label">電気料金（BESS なし比）</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">74%</span>
  <span class="kpi-label">PV 自家消費率</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">0.4 s</span>
  <span class="kpi-label">求解時間（48 ステップ）</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">1.2 回</span>
  <span class="kpi-label">等価サイクル数/日</span>
</div>

</div>

---

<!-- _class: pros-cons -->

# 考察｜有効性と限界

<div class="pc-pros">
<li>厳密最適解を 1 秒未満で取得 — 組込み機器でも実行可能</li>
<li>価格体系の変更は係数の差し替えのみで追従</li>
<li>ルールベース比 17% の料金削減を一貫して達成</li>
</div>

<div class="pc-cons">
<li>完全予見を仮定 — 予測誤差下の性能は別途検証が必要</li>
<li>蓄電池劣化コストを陽に最小化していない</li>
<li>需要家間の電力融通（P2P）は対象外</li>
</div>

---

<!-- _class: takeaway -->

# まとめ

<div class="ta-main">充放電スケジューリングは LP に帰着でき、厳密最適が実用時間で手に入る</div>

<div class="ta-points">
<li>売買分解により双方向取引を線形のまま定式化（決定変数 192・制約 300）</li>
<li>夏季 31 日間でルールベース比 17%・BESS なし比 28% の料金削減</li>
<li>今後: 予測誤差を考慮したロバスト化と劣化コストの内生化</li>
</div>

---

<!-- _class: references -->

# 参考文献

<ol>
<li>
  <span class="author">著者名</span>
  <span class="title">"論文タイトル."</span>
  <span class="venue">ジャーナル名, vol. X, no. Y, 2026.</span>
</li>
<li>
  <span class="author">著者名</span>
  <span class="title">"論文タイトル."</span>
  <span class="venue">国際会議名, 2025.</span>
</li>
<li>
  <span class="author">著者名</span>
  <span class="title">"論文タイトル."</span>
  <span class="venue">電気学会論文誌B, 2024.</span>
</li>
<li>
  <span class="author">著者名</span>
  <span class="title">"書籍・技術報告のタイトル."</span>
  <span class="venue">出版社／機関, 2023.</span>
</li>
<li>
  <span class="author">著者名</span>
  <span class="title">"データセット・ツールの名称."</span>
  <span class="venue">URL または DOI, 参照 2026-XX-XX.</span>
</li>
</ol>

---

<!-- _class: end -->

# ご清聴ありがとうございました

ご質問をお願いいたします

your.name@example.ac.jp
