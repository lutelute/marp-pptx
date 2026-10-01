---
marp: true
theme: academic
paginate: true
---

<!-- _class: survey -->
<!-- _mark: attention -->

# 関連研究｜注意機構の計算量を下げる 3 系統

## 共通の課題｜どの系統も、精度・計算量・実装の手間のどれかを手放している

<div class="sv-group">
<span class="sv-label">**疎な注意**パターン</span>
<span class="sv-gap">固定パターンの外にある長距離の依存を取りこぼしうる</span>

- R. Child et al., "Generating long sequences with sparse transformers," arXiv:1904.10509, 2019.
- I. Beltagy et al., "Longformer: The long-document transformer," arXiv:2004.05150, 2020.
- M. Zaheer et al., "Big Bird: Transformers for longer sequences," NeurIPS, 2020.
</div>

<div class="sv-group">
<span class="sv-label">**低ランク・カーネル**近似</span>
<span class="sv-gap">近似の誤差がそのまま精度の低下になる</span>

- S. Wang et al., "Linformer: Self-attention with linear complexity," arXiv:2006.04768, 2020.
- K. Choromanski et al., "Rethinking attention with performers," ICLR, 2021.
</div>

<div class="sv-group">
<span class="sv-label">**IO を意識した**厳密計算</span>
<span class="sv-gap">速く省メモリになるが、計算量のオーダーは変わらない</span>

- T. Dao et al., "FlashAttention: Fast and memory-efficient exact attention with IO-awareness," NeurIPS, 2022.
</div>

<div class="sv-verdict">未解決: 精度を保ったまま、長い系列で計算量そのものを下げる方法</div>
