---
marp: true
---

<!-- _class: title -->

# バグ再現チェック：スキル文書どおりの書き方
## — 修正後はすべて正しく描画されるはず —

重信 竜人 $^{1}$, 検証 太郎 $^{2}$

$^{1}$ 福井大学 工学研究科 &emsp; $^{2}$ 検証研究所

Regression Check 2026 &ensp;|&ensp; 2026年6月

---

<!-- _class: timeline-h -->

# timeline-h：detail 内の br タグ

<div class="tl-h-container">

<div class="tl-h-item">
  <div class="tl-h-block">
    <span class="tl-h-year">2012</span>
    <span class="tl-h-text">AlexNet</span>
    <div class="tl-h-detail">CNN + ImageNet<br>画像認識の躍進</div>
  </div>
</div>

<div class="tl-h-item">
  <div class="tl-h-block">
    <span class="tl-h-year">2017</span>
    <span class="tl-h-text">Transformer</span>
    <div class="tl-h-detail">Self-attention<br>並列処理革命</div>
  </div>
</div>

<div class="tl-h-item highlight">
  <div class="tl-h-block">
    <span class="tl-h-year">2026</span>
    <span class="tl-h-text bold">本研究</span>
    <div class="tl-h-detail">Sparse + 収束保証<br>計算量60%削減</div>
  </div>
</div>

</div>

---

<!-- _class: takeaway -->

# takeaway：長い主文の折り返し

<div class="ta-main">感度行列で「ほぼ正しい解」を一瞬で出し、AC 潮流で「正しさ」を数回だけ買い足す — この分業が実用速度と工学精度を両立させる</div>

<div class="ta-points">
<li>LP 化により 1,200 ノードで 47 倍の高速化、誤差は ±1.8% に収束</li>
<li>感度はヤコビアンの再利用で得られ、既存潮流計算コードに後付け可能</li>
<li>系統構成変更にも再線形化のみで追従し、ML 代理モデルの再学習コストを回避</li>
</div>

---

<!-- _class: table-slide -->

# table-slide：6行表と box-accent の間隔

| 手法 | 受入量誤差 | 計算時間 | 系統構成変更への追従 |
|------|:--------:|:--------:|:------------------:|
| 総当たり AC（基準） | — | 94.1 分 | 再計算が必要 |
| モンテカルロ確率評価 | ±2.1% | 38.5 分 | 再サンプリング必要 |
| LinDistFlow 線形化 | ±5.8% | 0.3 分 | 即時 |
| ML 代理モデル | ±3.4% | 0.1 分 | **再学習が必要** |
| **提案手法** | **±1.8%** | **2.0 分** | **即時（再線形化のみ）** |

<div class="box-accent">

**位置づけ**: 表の最下行とこのボックスの間に、きちんと余白が確保されているかを確認する。

</div>

<div class="footnote">修正後: 表ブロックに +0.25in のヘッドルームが入る。</div>
