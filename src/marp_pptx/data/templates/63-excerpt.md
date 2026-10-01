---
marp: true
theme: academic
paginate: true
---

<!-- _class: excerpt -->

# 研究課題（補足）｜DER の aggregation

## 他アプローチ｜フレキシビリティを集約して扱う（Zhou et al., 2025）

<div class="ex-item">
<span class="ex-quote">"… directly incorporating a large population of DERs in system-wide scheduling introduces high computing efforts."</span>
<span class="ex-read">DER を個別に直接制御すると==計算負荷が大きい==</span>
</div>

<div class="ex-item">
<span class="ex-quote">"Another promising solution is DER aggregation, which requires the identification of the Aggregated Feasible Power Region (AFPR) … while respecting the network constraints."</span>
<span class="ex-read">集約した実行可能電力領域 AFPR（＝ネットワーク制約を充足する領域）を同定する</span>
</div>

<div class="ex-verdict">**利点**　ネットワーク制約を満たしたまま TSO-DSO 協調に使える<br>**課題**　==不確実性下==での制約充足と disaggregation（計算負荷が大きい）</div>

<div class="ex-source">Y. Zhou, C. Essayeh and T. Morstyn, "Aggregated feasible active power region for distributed energy resources with a distributionally robust joint probabilistic guarantee," IEEE Trans. Power Syst., vol. 40, no. 1, pp. 556–571, 2025.</div>
