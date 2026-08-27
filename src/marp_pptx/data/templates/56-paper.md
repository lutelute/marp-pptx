---
marp: true
theme: academic
paginate: true
---

<!-- _class: paper -->

# Attention Is All You Need

## 文献報告｜再帰を捨てた自己注意のみの系列変換

<div class="pp-meta">
  <span class="pp-authors">Vaswani, A., Shazeer, N., Parmar, N. et al.（Google Brain / Google Research）</span>
  <span class="pp-venue">NeurIPS 2017</span>
  <span class="pp-stats">被引用 130,000+ ／ arXiv:1706.03762</span>
</div>

<div class="pp-why">
  <span class="pp-why-label">選定理由</span>
  <span class="pp-why-body">自研究のスパース注意機構の出発点。計算量 $O(n^2)$ がどの設計判断から生じたかを原典で確認し、削減余地を特定する。</span>
</div>

<div class="pp-points">
- 再帰・畳み込みを排し **Multi-Head Self-Attention** のみで系列変換を構成
- 位置情報は正弦波の位置エンコーディングで注入（学習不要）
- WMT14 EN-DE で BLEU ==28.4==（当時 SOTA）、学習コストは従来比 1/4
</div>
