---
marp: true
---

<!-- _class: title -->

# 数式注釈とステップ強調が PPTX に来た
## 新機能ギャラリー — Phase 1

marp-pptx × marp-theme-tmu-cs（記法互換）

2026-06-12 / docs/FEATURE-DESIGN.md

<!-- note: tmu-cs の [!math-annotate] と [!step] 記法をそのまま採用。同じ MD が tmu-cs(HTML/PDF) でも marp-pptx(編集可能PPTX) でも意味を保つ。 -->

---

<!-- _class: statement -->

tmu-cs の記法が、そのまま編集可能 PPTX になる。

---

<!-- _class: equation -->

# 数式注釈｜行末コメントが注釈カードになる

<div class="eq-main">

$$
X_k % [!math-annotate label="出力" note="k 番目の周波数成分"]
= \sum_{n=0}^{N-1} % [!math-annotate note="全サンプルにわたる総和"]
x_n\, e^{-i 2\pi k n / N} % [!math-annotate label="信号" note="離散時間信号 × 回転因子"]
$$

</div>

<div class="footnote">記法は marp-theme-tmu-cs の % [!math-annotate label= note= color=] と互換。注釈カードとコネクタは実シェイプなので PowerPoint で動かせる。</div>

---

<!-- _class: equation -->

# 数式注釈｜color 指定で系統を塗り分ける

<div class="eq-main">

$$
\max_{p \ge 0}\ \sum_{i \in \mathcal{G}} p_i % [!math-annotate label="目的" note="PV 連系量の合計を最大化する"]
\text{s.t.}\quad V_{\min} \le V^{(0)} + S\,p \le V_{\max} % [!math-annotate label="制約" note="全ノードの電圧を適正範囲に保つ" color="#0d558e"]
S = \partial V / \partial p % [!math-annotate note="感度行列はヤコビアンの再利用で得る" color="#b3502f"]
$$

</div>

<div class="footnote">ホスティングキャパシティ評価の LP（demo/hosting-capacity-v3.md）。color="#hex" がカードのラベルとコネクタに効く。</div>

---

<!-- _class: code -->

# ステップ強調｜スライドが自動で 3 枚に展開される

<div class="cd-code">

```python
S = jacobian_from_powerflow(net)    # 基準潮流から感度を取る [!step 1 info]
lp = build_lp(S, v_limits)          # [!step 2 highlight:2]
p_opt = lp.solve()
violations = ac_verify(p_opt)       # 違反があれば再線形化 [!step 3 warning]
print(f"HC = {sum(p_opt):.1f} kW")
```

</div>

<div class="cd-desc">行末の [!step N action[:M]] がステップ。クリック送りで 1→2→3 と強調が移る（tmu-cs の段階表示と同じ「動き」）。:2 のような span 指定で複数行をまとめて照らす。</div>

---

<!-- _class: takeaway -->

# この記法は両方で動く

<div class="ta-main">Marp（HTML/PDF）でも marp-pptx（編集可能 PPTX）でも、同じ MD が同じ意味を持つ</div>

<div class="ta-points">
<li>% [!math-annotate label= note= color=] — 数式行に注釈カード＋コネクタ</li>
<li># [!step N action[:M]] — 行強調＋ステップごとのスライド自動展開（=動き）</li>
<li>Phase 2 は bib 引用 [@key]・セクション/TOC 自動生成（docs/FEATURE-DESIGN.md）</li>
</div>

---

<!-- _class: end -->

# 機能ギャラリーはここまで

marp-pptx convert demo/feature-gallery.md -p tmu-cs

Phase 2 へ続く
