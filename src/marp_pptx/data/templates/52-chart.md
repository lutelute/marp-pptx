---
marp: true
theme: claude
---

<!-- _class: chart -->
<!-- _chart: column -->

# 系列長ごとの計算時間（相対）

| 系列長 | Transformer | Ours |
|---|---|---|
| 1024 | 1.0 | 0.6 |
| 4096 | 4.2 | 1.1 |
| 16384 | 18.5 | 2.3 |

<div class="chart-caption">同一ハードウェアで測定。Ours は Okabe-Ito 安全色で強調。</div>
