---
marp: true
theme: academic
paginate: true
---

<!-- _class: survey -->
<!-- _mark: distributionally robust -->

# 関連研究｜ネットワーク制約の扱い

## 共通の短所｜緩和・近似モデルではネットワーク制約の充足が保証されない

<div class="sv-group">
<span class="sv-label">電力潮流方程式の**線形近似**モデル</span>
<span class="sv-gap">近似誤差の分だけ電圧・電流の制約を逸脱しうる</span>

- Y. Wen et al., "Centralized distributionally robust chance-constrained dispatch of integrated transmission-distribution systems," IEEE Trans. Power Syst., 2024.
- S. Talari et al., "Sequential clearing of network-aware local energy and flexibility markets in community-based grids," IEEE Trans. Smart Grid, 2024.
- X. Shi et al., "Day-ahead distributionally robust optimization-based scheduling for distribution systems with electric vehicles," IEEE Trans. Smart Grid, 2023.
- Y. Zhou et al., "Aggregated feasible active power region for distributed energy resources with a distributionally robust joint probabilistic guarantee," IEEE Trans. Power Syst., 2025.
</div>

<div class="sv-group">
<span class="sv-label">電力潮流方程式の**二次錐緩和**モデル</span>
<span class="sv-gap">緩和が厳密にならない条件では解が物理的に実現できない</span>

- T. Jiang et al., "Flexibility clearing in joint energy and flexibility markets considering TSO-DSO coordination," IEEE Trans. Smart Grid, 2023.
</div>

<div class="sv-group">
<span class="sv-label">アフィン方策で**緩和**＋機会制約を**近似**</span>

- M. Rayati et al., "Distributionally robust chance constrained optimization for providing flexibility in an active distribution network," IEEE Trans. Smart Grid, 2022.
</div>

<div class="sv-verdict">未解決: 不確実性の下でネットワーク制約の充足を保証しつつ、計算負荷を抑える方法</div>
