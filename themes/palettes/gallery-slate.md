---
marp: true
theme: academic-slate
paginate: true
math: katex
---

<!-- _class: title -->
<!-- _paginate: false -->

# Slate — ウルトラミニマル
## テンプレートギャラリー

青灰。装飾なし。最大余白。静かな発表に。

v1.0

---

<!-- _class: agenda -->

# [Agenda] 目次

<div class="agenda-list">

1. 背景と課題
2. 提案手法
3. 実験結果
4. まとめ

</div>

---

<!-- _class: divider -->
<!-- _paginate: false -->

# セクション区切り

## サブタイトル

---

# [Default] 本文テンプレート

## 問題設定

- 背景情報を記述する
- 既存手法の課題を明示する
- 本研究の位置づけを示す

## 貢献

<div class="box-accent">

1. 貢献 1: 計算量を **60% 削減**
2. 貢献 2: 理論的な収束保証
3. 貢献 3: 3ベンチマークで SOTA

</div>

---

<!-- _class: equation -->

# [Equation] 数式テンプレート

<div class="eq-main">

$$\mathcal{L}(\theta) = -\frac{1}{N}\sum_{i=1}^{N} \log p(y_i | x_i; \theta)$$

</div>

<div class="eq-desc">
  <span class="sym">$\theta$</span>
  <span>モデルパラメータ</span>
  <span class="sym">$N$</span>
  <span>サンプル数</span>
</div>

---

<!-- _class: sandwich -->

# [Sandwich] 上下挟みカラム

<div class="top">
<p class="lead">リード文で全体の文脈を説明する。</p>
</div>

<div class="columns c3">
<div>

### A

- 項目 1
- 項目 2

</div>
<div>

### B

- 項目 1
- 項目 2

</div>
<div>

### C

- 項目 1
- 項目 2

</div>
</div>

<div class="bottom">
<div class="conclusion">

**まとめ**: 要点を1文で。

</div>
</div>

---

<!-- _class: zone-flow -->

# [Zone-Flow] フロー

<div class="zf-container">
<div class="zf-box">
  <span class="zf-label">Input</span>
  <span class="zf-body">データ収集と前処理</span>
</div>
<div class="zf-box">
  <span class="zf-label">Process</span>
  <span class="zf-body">モデル学習と最適化</span>
</div>
<div class="zf-box">
  <span class="zf-label">Output</span>
  <span class="zf-body">評価と比較</span>
</div>
</div>

---

<!-- _class: zone-compare -->

# [Zone-Compare] 比較

<div class="zc-container">
<div class="zc-left">
  <span class="zc-label">従来手法</span>
  <span class="zc-body">計算量 $O(n^2)$。精度は高いがスケーラビリティに課題。</span>
</div>
<div class="zc-vs">VS</div>
<div class="zc-right">
  <span class="zc-label">提案手法</span>
  <span class="zc-body">$O(n\sqrt{n})$ で同等精度。GPU 並列化対応。</span>
</div>
</div>

---

<!-- _class: summary -->

# [Summary] まとめ

<ol class="summary-points">
<li>成果 1</li>
<li>成果 2</li>
<li>成果 3</li>
</ol>

---

<!-- _class: end -->
<!-- _paginate: false -->

# Thank you

Questions?
