---
marp: true
---

<!-- _class: title -->

# 勾配法の収束解析と実装
## beamer スタイルデモ — 定理ブロック・フッターバー・ステップ強調

山田 太郎 — ○○大学大学院

研究会発表 / 2026

<!-- note: -p beamer のデモ。紺タイトル帯・フッターバー(タイトル|セクション|頁)・blocks 型(定理環境)・math-annotate・code step を全部載せ。 -->

---

<!-- _class: agenda -->

# Outline

<div class="agenda-list">

1. 理論 — 強凸・平滑関数に対する収束性の定理
2. アルゴリズム — 10 行の実装とステップ実行
3. 評価 — 条件数別の反復回数を実測と比較
4. まとめ — レートの読み方と次の一手

</div>

---

<!-- _class: divider -->

# 1. 理論
## 収束性の定理と証明のスケッチ

---

<!-- _class: blocks -->

# 主結果

<div class="bk-container">

<div class="bk theorem">
  <span class="bk-title">定理 1（線形収束）</span>
  <span class="bk-body">$f$ が $\mu$-強凸かつ $L$-平滑のとき、ステップ幅 $\eta = 1/L$ の勾配法は $\|x_{k+1} - x^*\| \le \rho\,\|x_k - x^*\|$（$\rho = 1 - \mu/L$）を満たす。</span>
</div>

<div class="bk example">
  <span class="bk-title">例（二次関数）</span>
  <span class="bk-body">$f(x) = \frac{1}{2}x^\top A x$ では $\mu, L$ は $A$ の最小・最大固有値。条件数 $\kappa = L/\mu$ が小さいほど速い。</span>
</div>

<div class="bk alert">
  <span class="bk-title">注意</span>
  <span class="bk-body">$\eta > 2/L$ では発散する。実務ではラインサーチか $L$ の上界推定を併用する。</span>
</div>

</div>

---

<!-- _class: equation -->

# 収束レートの構造

<div class="eq-main">

$$
\rho % [!math-annotate label="レート" note="1 反復ごとの誤差縮小率"]
= 1 - % [!math-annotate note="1 から離れるほど速い"]
\frac{\mu}{L} % [!math-annotate label="条件数の逆数" note="問題の曲がり具合の比" color="#9f1d1d"]
$$

</div>

<div class="footnote">tmu-cs 互換の [!math-annotate] — beamer テーマでも同じ記法がそのまま効く。</div>

---

<!-- _class: divider -->

# 2. アルゴリズム
## 実装と計算手順

---

<!-- _class: code -->

# 実装は 10 行で済む

<div class="cd-code">

```python
def gradient_descent(grad, x0, L, tol=1e-8):
    x = x0                          # 初期点 [!step 1 info]
    eta = 1.0 / L                   # 定理 1 のステップ幅 [!step 2 highlight]
    while True:
        g = grad(x)                 # 勾配を評価 [!step 3 focus:2]
        x_new = x - eta * g
        if abs(x_new - x) < tol:    # 収束判定 [!step 4 warning]
            return x_new
        x = x_new
```

</div>

<div class="cd-desc">クリック送りでステップが進む。フッターバーには現在のセクション「アルゴリズム」が出ている。</div>

---

<!-- _class: divider -->

# 3. 評価
## 数値実験

---

<!-- _class: table-slide -->

# 条件数別の反復回数

| 条件数 $\kappa$ | 理論上界 | 実測 |
|---:|---:|---:|
| 10 | 23 回 | 19 回 |
| 100 | 230 回 | 187 回 |
| 1000 | 2,303 回 | 1,956 回 |

<div class="footnote">許容誤差 $10^{-8}$、ランダム二次関数 100 試行の中央値。理論上界 $\kappa \ln(1/\varepsilon)$ と整合。</div>

---

<!-- _class: takeaway -->

# まとめ

<div class="ta-main">強凸＋平滑なら、勾配法は条件数で決まる線形レートで必ず収束する</div>

<div class="ta-points">
<li>定理 1 のレート $\rho = 1 - \mu/L$ は実測と整合（上界の 8 割程度）</li>
<li>実装は 10 行 — ステップ幅は $1/L$ にするだけ</li>
<li>条件数が大きい問題は前処理か加速法（Nesterov）を検討</li>
</div>

---

<!-- _class: end -->

# ご清聴ありがとうございました

このデッキ: marp-pptx convert demo/beamer-demo.md -p beamer
