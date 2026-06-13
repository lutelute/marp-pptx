---
marp: true
---

<!-- _class: title -->

# 近接勾配法による制約付き最適化の収束解析
## 強凸性を仮定しない O(1/k) レートと加速

山田 太郎¹, 佐藤 花子¹, 鈴木 一郎²

¹○○大学大学院 工学研究科 / ²△△研究所 / 最適化研究会 2026

<!-- note: 高密度アカデミック版デモ。-p beamer。タイトルは Madrid 流の紺角丸ボックス、各フレームは紺帯タイトル。 -->

---

<!-- _class: agenda -->

# Outline

<div class="agenda-list">

1. 背景と問題設定 — 合成最適化と近接作用素、既存手法の限界
2. 理論 — 主定理 O(1/k) と証明スケッチ、加速版 O(1/k²) への接続
3. アルゴリズムと数値評価 — 12 行の実装と LASSO ベンチマーク
4. まとめと今後 — 非凸拡張と確率的設定への展望

</div>

---

<!-- _class: divider -->

# 1. 背景と問題設定
## 合成最適化問題と近接勾配法

---

<!-- _class: sandwich -->

# 合成最適化｜滑らかな損失＋非平滑な正則化

<div class="top">

<p class="lead">機械学習・信号処理の推定問題は $\min_x f(x) + g(x)$（$f$: 滑らかな損失、$g$: 非平滑な正則化）に帰着する。$g$ の非平滑性が単純な勾配法を阻む一方、$g$ の近接作用素は多くの場合閉形式で計算できる。</p>

</div>

<div class="columns c2">
<div>

## 代表例

- **LASSO**: $f = \frac{1}{2}\|Ax-b\|^2$, $g = \lambda\|x\|_1$
- **画像復元**: $f$ = データ忠実項, $g$ = 全変動 (TV)
- **行列補完**: $g$ = 核ノルム（低ランク誘導）
- 共通点: $\mathrm{prox}_{\eta g}$ が**閉形式**（軟しきい値等）

</div>
<div>

## 既存手法の限界

- 劣勾配法: $O(1/\sqrt{k})$ と遅く、ステップ調整が難しい
- 平滑化: 近似誤差と速度のトレードオフが残る
- 内点法: 高精度だが次元 10⁵ 超の大規模問題で破綻
- 強凸仮定の解析が多いが、LASSO の $f$ は**強凸でない**

</div>
</div>

<div class="bottom">

<div class="conclusion">

**本研究**: 強凸性なしで近接勾配法の $O(1/k)$ を初等的な単調性議論だけで示し、同じ枠組みで Nesterov 加速の $O(1/k^2)$ まで一望する。

</div>

</div>

---

<!-- _class: cols-2 -->

# 問題設定｜仮定と記号

<div class="columns">
<div>

## 仮定

- (A1) $f: \mathbb{R}^n \to \mathbb{R}$ は凸かつ $L$-平滑（$\nabla f$ が $L$-リプシッツ）
- (A2) $g: \mathbb{R}^n \to (-\infty, +\infty]$ は凸・下半連続（非平滑可）
- (A3) 最適解集合 $X^* \neq \emptyset$、最適値 $F^* > -\infty$
- (A4) $\mathrm{prox}_{\eta g}$ は厳密に計算可能
- 強凸性 ($\mu > 0$) は**仮定しない**

</div>
<div>

## 記号

- $F(x) = f(x) + g(x)$ — 目的関数
- $\mathrm{prox}_{\eta g}(v) = \arg\min_u \left( g(u) + \frac{1}{2\eta}\|u - v\|^2 \right)$
- $x^+ = \mathrm{prox}_{\eta g}(x - \eta \nabla f(x))$ — 1 反復
- $G_\eta(x) = (x - x^+)/\eta$ — 勾配写像（停留性の尺度）
- $R_0 = \|x_0 - x^*\|$ — 初期距離

</div>
</div>

<div class="footnote">表記は Beck (2017), "First-Order Methods in Optimization" に準拠。$\eta$ はステップ幅で、本解析では固定 $\eta = 1/L$。</div>

---

<!-- _class: divider -->

# 2. 理論
## 主定理と証明のスケッチ

---

<!-- _class: blocks -->

# 主結果｜強凸性なしの大域収束レート

<div class="bk-container">

<div class="bk theorem">
  <span class="bk-title">定理 1（$O(1/k)$ 収束 — 本研究）</span>
  <span class="bk-body">仮定 (A1)–(A4) のもと、$\eta = 1/L$ の近接勾配法は任意の $x_0$ から $F(x_k) - F^* \le \dfrac{L \|x_0 - x^*\|^2}{2k}$ を満たす。さらに列 $\{F(x_k)\}$ は単調非増加で、$\{x_k\}$ はある $x^* \in X^*$ に収束する。</span>
</div>

<div class="bk lemma">
  <span class="bk-title">補題 2（十分減少 + 距離単調性）</span>
  <span class="bk-body">1 反復で $F(x^+) \le F(x) - \frac{\eta}{2}\|G_\eta(x)\|^2$ かつ $\|x^+ - x^*\| \le \|x - x^*\|$。この 2 つの単調性だけで定理 1 が従う（テレスコープ和）。</span>
</div>

<div class="bk alert">
  <span class="bk-title">注意（レートの最適性）</span>
  <span class="bk-body">$O(1/k)$ は近接勾配法に対してタイト（下界が存在）。これより速くするには加速（運動量）が必須 — 次頁の FISTA で $O(1/k^2)$。</span>
</div>

</div>

---

<!-- _class: cols-2 -->

# 証明のスケッチ｜5 ステップで閉じる

<div class="columns">
<div>

## 前半: 1 反復の評価

1. $f$ の $L$-平滑性から降下補題: $f(x^+) \le f(x) + \langle \nabla f(x), x^+ - x \rangle + \frac{L}{2}\|x^+ - x\|^2$
2. $\mathrm{prox}$ の最適性条件（変分不等式）を $g$ に適用
3. 1+2 を足すと任意の $u$ に対し $F(x^+) \le F(u) + \frac{L}{2}\left(\|x - u\|^2 - \|x^+ - u\|^2\right)$

</div>
<div>

## 後半: 和をとる

4. $u = x^*$ で評価し $k$ 反復分を**テレスコープ和** — 距離項が打ち消し合い $\sum_k (F(x_k) - F^*) \le \frac{L}{2} R_0^2$
5. $F(x_k)$ の単調性（補題 2）で平均を最終値に置換 → $F(x_k) - F^* \le \frac{L R_0^2}{2k}$ ∎

</div>
</div>

<div class="footnote">鍵は手順 3 の「3 点不等式」。強凸性は一切使わず、距離のテレスコープだけで閉じるのが本証明の特徴。加速版は手順 4 の重み付けを $t_k^2$ に変えるだけ。</div>

---

<!-- _class: equation -->

# 加速の構造｜FISTA のレートを項ごとに読む

<div class="eq-main">

$$
F(x_k) - F^* % [!math-annotate label="最適性ギャップ" note="k 反復後の目的関数値の差"]
\;\le\; \frac{2L}{(k+1)^2} % [!math-annotate label="レート" note="運動量により 1/k が 1/k² に加速" color="#9f1d1d"]
\, \|x_0 - x^*\|^2 % [!math-annotate label="初期距離" note="スタート地点の良さがそのまま定数に" color="#14653d"]
$$

</div>

<div class="footnote">FISTA (Beck & Teboulle 2009)。係数 2L/(k+1)² は外挿パラメータ $t_{k+1} = (1+\sqrt{1+4t_k^2})/2$ から。注釈は [!math-annotate] 記法。</div>

---

<!-- _class: divider -->

# 3. アルゴリズムと数値評価
## 実装・比較実験

---

<!-- _class: code -->

# 実装｜近接勾配法は 8 行、FISTA でも +4 行

<div class="cd-code">

```python
def fista(grad_f, prox_g, x0, L, K):
    x = y = x0; t = 1.0
    eta = 1.0 / L                       # 定理 1 と同じステップ幅 [!step 1 info]
    for k in range(K):
        v = y - eta * grad_f(y)         # 勾配ステップ [!step 2 highlight]
        x_new = prox_g(v, eta)          # 近接ステップ（軟しきい値等）[!step 2 highlight:1]
        t_new = (1 + (1 + 4 * t * t) ** 0.5) / 2
        y = x_new + ((t - 1) / t_new) * (x_new - x)   # 運動量（外挿）[!step 3 warning]
        x, t = x_new, t_new
    return x
```

</div>

<div class="cd-desc">[!step] のクリック送りで 1→2→3 と強調が移る。運動量の 2 行を消せばそのまま定理 1 の近接勾配法 — 1 反復の計算量（勾配 1 回＋prox 1 回）は同一。</div>

---

<!-- _class: table-slide -->

# LASSO ベンチマーク｜目標精度到達までの反復数

| 設定（$n$ × 疎度） | 劣勾配法 | 近接勾配（定理 1） | FISTA | 高速化倍率 |
|---|---:|---:|---:|---:|
| $10^4$ × 5% | 41,200 | 1,840 | 214 | ×193 |
| $10^5$ × 1% | >10⁵（未達） | 6,310 | 521 | — |
| $10^5$ × 5%・相関設計 | >10⁵（未達） | 18,400 | 1,260 | — |

<div class="box-accent">

**読み方**: 理論レート（1/√k → 1/k → 1/k²）の差がそのまま反復数の桁差に現れる。相関設計では $L$ が増大し全手法が遅化するが、序列は不変。

</div>

<div class="footnote">目標 $F(x_k) - F^* \le 10^{-6}$、$\lambda = 0.1\|A^\top b\|_\infty$、乱数 10 シードの中央値。単一 CPU スレッド、warm start なし。</div>

---

<!-- _class: pros-cons -->

# 考察｜単調性ベース解析の射程と限界

<div class="pc-pros">
<li>証明が初等的（降下補題＋3 点不等式のみ）— 学部講義・教科書に載せられる構成</li>
<li>距離単調性は実装検証にも使える: $\|x_k - x^*\|$ が増えたらバグか $\eta$ 過大</li>
<li>同じ骨格が射影勾配・座標降下・ADMM の一部に流用可能</li>
<li>加速版へは和の重み付け変更のみで接続（統一的な見通し）</li>
</div>

<div class="pc-cons">
<li>非凸 $f$ には距離単調性が崩れ、別の Lyapunov 関数が必要</li>
<li>確率的勾配（ミニバッチ）では分散項が残り、このままでは $O(1/k)$ が出ない</li>
<li>$\mathrm{prox}$ が閉形式でない $g$（重複群正則化等）は内部反復の誤差解析が別途要る</li>
<li>$L$ が未知の実問題ではバックトラッキング直線探索の併用が前提になる</li>
</div>

---

<!-- _class: takeaway -->

# まとめ

<div class="ta-main">強凸性がなくても、2 つの単調性（値の減少・距離の非増加）だけで近接勾配法の $O(1/k)$ は閉じる</div>

<div class="ta-points">
<li>主定理: $F(x_k) - F^* \le L R_0^2 / 2k$（定理 1）。証明は降下補題＋3 点不等式＋テレスコープの 5 手</li>
<li>実証: LASSO で劣勾配法の 20〜200 倍速、FISTA でさらに 1 桁（理論序列どおり）</li>
<li>今後: 非凸拡張（KL 性に基づく Lyapunov 構成）と確率的設定（分散縮小との接続）</li>
</div>

---

<!-- _class: end -->

# ご清聴ありがとうございました

このデッキ: marp-pptx convert demo/beamer-demo.md -p beamer
