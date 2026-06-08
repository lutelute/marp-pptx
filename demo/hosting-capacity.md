---
marp: true
---

<!-- _class: title -->

# 配電系統における PV ホスティングキャパシティの高速評価
## — 電圧感度行列による線形化と AC 潮流逐次補正 —

重信 竜人 — 福井大学 工学研究科

電力技術研究会 2026 / 2026年6月（品質検証デモ）

<!-- note: 本デッキは marp-pptx の品質検証用デモ。数値は架空だがドメイン的に妥当な設定。 -->

---

<!-- _class: agenda -->

# 本日の内容

<div class="agenda-list">

1. 背景 — PV 大量連系と電圧制約
2. 研究課題
3. 提案手法 — 感度行列 × 逐次補正
4. 数値実験 — 1,200 ノード配電系統
5. まとめ

</div>

---

<!-- _class: rq -->

# Research Question

<div class="rq-main">
配電系統の PV 受入可能量（ホスティングキャパシティ）を、AC 潮流並みの精度を保ったまま実用時間で評価できるか？
</div>

<div class="rq-sub">
— 電圧感度行列による線形化と、AC 潮流による逐次補正の組み合わせ
</div>

---

<!-- _class: sandwich -->

# 背景：PV 連系申請の急増と電圧上昇制約

<div class="top">

<p class="lead">PV 連系申請は増加を続け、6.6 kV 配電フィーダでは適正電圧範囲（101±6 V 換算）の上限側逸脱が連系可否を決める支配的制約になっている。</p>

</div>

<div class="columns c2">
<div>

## 従来実務の課題

- 全ノード総当たりの AC 潮流計算で受入量を探索
- フィーダ規模 1,000 ノード超では**数時間オーダー**
- 申請件数の増加に計算が追いつかない

</div>
<div>

## 求められる要件

- 判定誤差は AC 潮流比で **数 % 以内**
- 1 フィーダあたり**数分以内**の応答
- 既存の系統データベースでそのまま動くこと

</div>
</div>

<div class="bottom">

<div class="conclusion">

**本研究の立場**: 感度行列による線形化で探索を高速化し、解の周辺だけ AC 潮流で逐次補正することで、速度と精度の両立を狙う。

</div>

</div>

---

<!-- _class: timeline-h -->

# ホスティングキャパシティ評価手法の変遷

<div class="tl-h-container">

<div class="tl-h-item">
  <div class="tl-h-block">
    <span class="tl-h-year">2012</span>
    <span class="tl-h-text">決定論的評価</span>
    <div class="tl-h-detail">最悪ケース想定で過度に保守的</div>
  </div>
</div>

<div class="tl-h-item">
  <div class="tl-h-block">
    <span class="tl-h-year">2016</span>
    <span class="tl-h-text">確率論的評価</span>
    <div class="tl-h-detail">モンテカルロ法、計算コスト大</div>
  </div>
</div>

<div class="tl-h-item">
  <div class="tl-h-block">
    <span class="tl-h-year">2019</span>
    <span class="tl-h-text">線形化潮流</span>
    <div class="tl-h-detail">LinDistFlow、軽負荷時に誤差</div>
  </div>
</div>

<div class="tl-h-item">
  <div class="tl-h-block">
    <span class="tl-h-year">2022</span>
    <span class="tl-h-text">ML 代理モデル</span>
    <div class="tl-h-detail">高速だが構成変更に弱い</div>
  </div>
</div>

<div class="tl-h-item highlight">
  <div class="tl-h-block">
    <span class="tl-h-year">2026</span>
    <span class="tl-h-text bold">本手法</span>
    <div class="tl-h-detail">感度行列＋逐次補正で両立</div>
  </div>
</div>

</div>

---

<!-- _class: divider -->

# 1. 提案手法

## 電圧感度行列による線形最適化と AC 潮流逐次補正

---

<!-- _class: equations -->

# 定式化：線形化ホスティングキャパシティ問題

<div class="eq-system">
  <div class="row">
    <span class="label">maximize</span>

$$\sum_{i \in \mathcal{G}} p_i$$

  </div>
  <div class="row">
    <span class="label">subject to</span>

$$V_{\min} \le V_j^{(0)} + (S\,p)_j \le V_{\max} \quad \forall j \in \mathcal{N}$$

  </div>
  <div class="row">
    <span class="label"></span>

$$\bigl| I_l^{(0)} + (T\,p)_l \bigr| \le I_l^{\mathrm{rate}} \quad \forall l \in \mathcal{L}$$

  </div>
  <div class="row">
    <span class="label"></span>

$$0 \le p_i \le \bar{p}_i$$

  </div>
</div>

<div class="eq-desc">
  <span class="sym">$p_i$</span>
  <span>ノード $i$ の PV 連系量（決定変数）</span>
  <span class="sym">$S_{ji} = \partial V_j / \partial p_i$</span>
  <span>電圧感度行列（基準潮流のヤコビアンから算出）</span>
  <span class="sym">$T_{li}$</span>
  <span>線路電流感度</span>
  <span class="sym">$V_j^{(0)},\ I_l^{(0)}$</span>
  <span>基準運転点の電圧・電流</span>
</div>

<div class="footnote">線形計画問題（LP）として解けるため大規模系統でも高速。感度は運転点近傍でのみ有効 → 次頁の逐次補正で誤差を抑える。</div>

<!-- note: 感度行列はニュートン法の最終ヤコビアンの逆行列から取り出すため、追加コストはほぼゼロである点を口頭で補足。 -->

---

<!-- _class: zone-process -->

# 解析パイプライン：LP と AC 潮流の往復

<div class="zp-container">

<div class="zp-step">
  <span class="zp-num">1</span>
  <span class="zp-title">基準潮流計算</span>
  <span class="zp-body">現行の負荷・発電条件で AC 潮流を 1 回解き、運転点とヤコビアンを取得。</span>
</div>

<div class="zp-step">
  <span class="zp-num">2</span>
  <span class="zp-title">感度行列の構築</span>
  <span class="zp-body">ヤコビアン逆行列から電圧・電流感度 $S, T$ を抽出。追加の潮流計算は不要。</span>
</div>

<div class="zp-step">
  <span class="zp-num">3</span>
  <span class="zp-title">線形最適化（LP）</span>
  <span class="zp-body">線形化制約のもとで受入量を最大化。1,200 ノードでも 1 秒未満で求解。</span>
</div>

<div class="zp-step">
  <span class="zp-num">4</span>
  <span class="zp-title">AC 検証と逐次補正</span>
  <span class="zp-body">LP 解で AC 潮流を再計算し、制約違反があれば運転点を更新して再線形化（平均 3 反復で収束）。</span>
</div>

</div>

---

<!-- _class: before-after -->

# 従来実務との比較：同一フィーダでの評価フロー

<div class="ba-before">
  <span class="ba-label">Before</span>
  <span class="ba-body">全ノード総当たりで連系量を二分探索し、その都度 AC 潮流を実行。1,200 ノードのモデルフィーダで約 94 分。申請が集中すると判定が翌日に持ち越し。</span>
</div>

<div class="ba-after">
  <span class="ba-label">After</span>
  <span class="ba-body">感度ベースの LP で一括最適化し、AC 補正は平均 3 回のみ。同一フィーダで約 2 分、最大電圧誤差 0.4%。申請当日中の回答が可能に。</span>
</div>

---

<!-- _class: chart -->
<!-- _chart: column -->

# 計算時間の比較（フィーダ規模別）

| ノード数 | 総当たり AC [分] | 提案手法 [分] |
|---|---|---|
| 300 | 6.2 | 0.4 |
| 600 | 24.8 | 0.9 |
| 1200 | 94.1 | 2.0 |

<div class="chart-caption">同一計算機（Xeon w5-2455X, 単スレッド）での実測。提案手法はノード数に対しほぼ線形にスケール。</div>

---

<!-- _class: kpi -->

# 数値実験の主要結果（1,200 ノード実フィーダ相当モデル）

<div class="kpi-container">

<div class="kpi-item">
  <span class="kpi-value">47×</span>
  <span class="kpi-label">高速化（94.1 分 → 2.0 分）</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">0.4%</span>
  <span class="kpi-label">最大電圧誤差（AC 潮流比）</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">3 回</span>
  <span class="kpi-label">平均 AC 補正反復数</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">98.2%</span>
  <span class="kpi-label">受入量の推定精度</span>
</div>

</div>

---

<!-- _class: table-slide -->

# 既存手法との比較（1,200 ノード・同一制約条件）

| 手法 | 受入量誤差 | 計算時間 | 系統構成変更への追従 |
|------|:--------:|:--------:|:------------------:|
| 総当たり AC（基準） | — | 94.1 分 | 再計算が必要 |
| モンテカルロ確率評価 | ±2.1% | 38.5 分 | 再サンプリング必要 |
| LinDistFlow 線形化 | ±5.8% | 0.3 分 | 即時 |
| ML 代理モデル | ±3.4% | 0.1 分 | **再学習が必要** |
| **提案手法** | **±1.8%** | **2.0 分** | **即時（再線形化のみ）** |

<div class="box-accent">

**位置づけ**: 純線形化の速度と AC 基準の精度の中間を取り、運用実務の「当日回答」要件を満たす唯一の構成。

</div>

<div class="footnote">誤差は真の受入量（総当たり AC）に対する相対値。時間は同一ハードウェアでの実測平均。</div>

---

<!-- _class: takeaway -->

# キーメッセージ

<div class="ta-main">速さは感度行列で、正しさは AC 潮流の逐次補正で買う</div>

<div class="ta-points">
<li>LP 化により 1,200 ノードで 47 倍の高速化、誤差は ±1.8% に収束</li>
<li>感度はヤコビアンの再利用で得られ、既存潮流計算コードに後付け可能</li>
<li>系統構成変更にも再線形化のみで追従し、ML 代理モデルの再学習コストを回避</li>
</div>

---

<!-- _class: end -->

# Thank you

Questions?

shigenobu@u-fukui.ac.jp
