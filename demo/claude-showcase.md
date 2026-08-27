---
marp: true
---

<!-- _class: title -->
# 判断の速さは、設計できる
## MARP-PPTX SHOWCASE
marp-pptx チーム / 2026

---

<!-- _class: split-panel -->

# 速さと正しさは、もう交換条件ではない

## PROPOSAL

<div class="sp-body">
- 感度行列 LP が **1 秒未満** で候補を出す
- AC 潮流は違反時のみ — 平均 ==3 往復== で収束
- 判定誤差は AC 比 ±1.8%、==当日回答== が標準になる
- 既存の系統データベースはそのまま使える
</div>

---

<!-- _class: big-number -->
<!-- source: 1,200ノード実測（Xeon w5-2455X） -->
# 到達した速度
<div class="big-number">
  <span class="bn-value">47×</span>
  <span class="bn-label">総当たり AC 比の高速化</span>
  <span class="bn-caption">94.1 分 → 2.0 分</span>
</div>

---

<!-- _class: flow -->

# 仕組みは 2 つの計算の分業

## 反復補正｜LP が候補を出し、AC が正しさを保証する

```mermaid
flowchart LR
  S[感度行列 S] --> LP[線形最適化 LP<br>求解 1 秒未満]:::accent
  LP -->|候補解| AC[AC 潮流計算<br>真値で検証]
  AC -->|違反なし| OUT([当日回答]):::primary
  AC -.->|違反あり| LP
```

---

<!-- _class: kpi -->
# 4 つの数字で言い切る
<div class="kpi-container">
<div><span class="kpi-value">47×</span><span class="kpi-label">高速化</span></div>
<div><span class="kpi-value">±1.8%</span><span class="kpi-label">誤差</span></div>
<div><span class="kpi-value">3 回</span><span class="kpi-label">平均反復</span></div>
<div><span class="kpi-value">0 件</span><span class="kpi-label">見逃し違反</span></div>
</div>

---

<!-- _class: dark -->

速さは感度行列で、正しさは AC 潮流で買う。

---

<!-- _class: end -->
# Thank you
marp-pptx — Markdown から、編集できる PowerPoint へ
