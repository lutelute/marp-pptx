---
marp: true
---

<!-- _class: title -->

# PV の受入可否を「当日」に答える
## 感度行列LP × AC潮流の分業

重信 竜人 — 福井大学 工学研究科

電力技術研究会 2026 / 品質検証デモ v2

<!-- note: v2: メッセージライン化 + 自作図解 + keynote 密度。数値は架空だがドメイン的に妥当な設定。 -->

---

<!-- _class: statement -->

「あと何 kW 繋げるか」への回答に、いま 94 分かかっている。

---

<!-- _class: diagram -->

# 電圧上限に当たった量が受入限界を決める

![w:880](assets/problem.png)

<div class="caption">PV の逆潮流で末端ほど電圧が上昇。上限に最初に触れた時点の合計連系量がホスティングキャパシティ（HC）。</div>

<!-- note: 従来は全ノード総当たりの AC 潮流でこの限界点を二分探索しており、1,200 ノードで 94 分かかる。 -->

---

<!-- _class: rq -->

# Research Question

<div class="rq-main">
AC の精度のまま、評価を分オーダーにできるか？
</div>

<div class="rq-sub">
— 速い近似と正確な検証を、それぞれ得意な側に分担させる
</div>

---

<!-- _class: diagram -->

# ほぼ正しい解は LP、正しさは AC で買い足す

![w:840](assets/loop.png)

<div class="caption">感度行列の LP が候補解を一瞬で出し、AC 潮流が真の電圧で検証する。違反時のみ再線形化 — 平均 3 往復で収束。</div>

---

<!-- _class: equation -->

# 感度行列で HC 評価はただの LP になる

<div class="eq-main">

$$\max_{p \ge 0}\ \sum_{i \in \mathcal{G}} p_i \qquad \text{s.t.} \quad V_{\min} \le V^{(0)} + S\,p \le V_{\max}$$

</div>

<div class="eq-desc">
  <span class="sym">$S = \partial V / \partial p$</span>
  <span>電圧感度行列（ヤコビアン再利用）</span>
  <span class="sym">$p$</span>
  <span>各ノードの PV 連系量（決定変数）</span>
</div>

<div class="footnote">線形計画なので 1,200 ノードでも求解 1 秒未満。線形化誤差は前頁の AC 往復で補正し、最終解は AC 検証済み。</div>

---

<!-- _class: chart -->
<!-- _chart: column -->

# 1,200 ノードで 94 分が 2 分になる

| ノード数 | 総当たり AC [分] | 提案手法 [分] |
|---|---|---|
| 300 | 6.2 | 0.4 |
| 600 | 24.8 | 0.9 |
| 1200 | 94.1 | 2.0 |

<div class="chart-caption">同一計算機・単スレッドでの実測。規模が大きいほど差が開く。</div>

---

<!-- _class: kpi -->

# 精度を落とさず 47 倍速い

<div class="kpi-container">

<div class="kpi-item">
  <span class="kpi-value">47×</span>
  <span class="kpi-label">高速化</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">1.8%</span>
  <span class="kpi-label">受入量の誤差（±）</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">0.4%</span>
  <span class="kpi-label">最大電圧誤差</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">3 回</span>
  <span class="kpi-label">平均 AC 補正</span>
</div>

</div>

---

<!-- _class: dark -->

速さは感度行列で、正しさは AC 潮流で買う。

---

<!-- _class: appendix -->

# 既存手法との定量比較

<span class="appendix-label">Appendix A</span>

| 手法 | 受入量誤差 | 計算時間 | 構成変更への追従 |
|------|:--------:|:--------:|:--------------:|
| 総当たり AC（基準） | — | 94.1 分 | 再計算 |
| モンテカルロ確率評価 | ±2.1% | 38.5 分 | 再サンプリング |
| LinDistFlow 線形化 | ±5.8% | 0.3 分 | 即時 |
| ML 代理モデル | ±3.4% | 0.1 分 | 再学習 |
| **提案手法** | **±1.8%** | **2.0 分** | **即時** |

---

<!-- _class: end -->

# Thank you

Questions?

shigenobu@u-fukui.ac.jp
