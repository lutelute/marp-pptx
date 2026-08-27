---
marp: true
theme: academic
paginate: true
---

<!-- _class: figure-story -->
<!-- source: Vaswani et al. (2017), Figure 1 -->

# アーキテクチャが並列化を解放する

## 読み方｜再帰が無いから、全時刻を同時に計算できる

![w:600](figures/architecture.png)

<div class="fs-points">
- **左**: エンコーダ — 自己注意＋FFN を $N=6$ 積層
- **右**: デコーダ — マスク付き自己注意で未来を遮蔽
- 系列方向の依存が無く、**全トークンを同時計算**
- 位置情報は正弦波の位置符号で注入（学習不要）
</div>

<div class="fs-conclusion">まとめ: 逐次計算の除去こそが本質 — 学習 3.5 日（8 GPU）での SOTA 到達は、この構造選択の直接の帰結。</div>
