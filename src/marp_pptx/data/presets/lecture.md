---
marp: true
---

<!-- _class: title -->
# 講義タイトル
## 第N回 — 今日のテーマ
担当教員 / 2026年度

---

<!-- _class: agenda -->
# 今日の流れ
<div class="agenda-list">
1. 用語の定義
2. 基本となる式
3. 計算の手順
4. まとめと演習
</div>

---

<!-- _class: definition -->
# 定義
<div class="df-term">フィードバック制御</div>
<div class="df-body">出力を測定し、目標値との差に基づいて入力を調整することで、外乱や変動があっても出力を目標に近づける制御方式。</div>
<div class="df-note">cf. フィードフォワード制御（事前補償）</div>

---

<!-- _class: equation -->
# 基本となる式（PID 制御）
<div class="eq-main">
$$u(t) = K_p e(t) + K_i \int_0^t e(\tau)\,d\tau + K_d \frac{de(t)}{dt}$$
</div>
<div class="eq-desc">
<span>$e(t)$</span><span>偏差（目標値 − 出力）</span>
<span>$K_p$</span><span>比例ゲイン</span>
<span>$K_i$</span><span>積分ゲイン</span>
<span>$K_d$</span><span>微分ゲイン</span>
</div>

---

<!-- _class: steps -->
# ゲイン調整の手順
<div class="st-container">
<div><span class="st-num">1</span><span class="st-title">P から</span><span class="st-body">$K_i=K_d=0$ で比例ゲインを上げる</span></div>
<div><span class="st-num">2</span><span class="st-title">I を追加</span><span class="st-body">定常偏差が消えるよう積分を足す</span></div>
<div><span class="st-num">3</span><span class="st-title">D で整える</span><span class="st-body">行き過ぎ（オーバーシュート）を抑える</span></div>
</div>

---

<!-- _class: summary -->
# まとめ
<ol class="summary-points">
<li>フィードバックは偏差に基づいて入力を調整する</li>
<li>PID は比例・積分・微分の3項からなる</li>
<li>調整は P → I → D の順が基本</li>
</ol>

---

<!-- _class: end -->
# 演習
次回までに章末問題 1〜3 を解いてくること。
