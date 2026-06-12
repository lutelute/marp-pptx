---
marp: true
theme: academic
paginate: true
math: katex
---

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
