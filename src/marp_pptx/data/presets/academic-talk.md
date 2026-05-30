---
marp: true
---

<!-- _class: title -->
# 研究タイトル
## サブタイトルがキッカーになる
山田 太郎 / 福井大学 / 2026年4月

---

<!-- _class: agenda -->
# 本日の内容
<div class="agenda-list">
1. 背景と研究目的
2. 提案手法
3. 実験結果
4. 考察とまとめ
</div>
<!-- note: 全体で15分。各章の時間配分に触れる -->

---

<!-- _class: rq -->
# 研究課題
<div class="rq-main">既存手法は大規模データに対してスケールするか？</div>
<div class="rq-sub">計算量 $O(n^2)$ がボトルネックとなっている。</div>

---

<!-- _class: sandwich -->
# 提案手法
<div class="top">
<div class="lead">従来の $O(n^2)$ を $O(n \log n)$ に改善。</div>
</div>
<div class="columns">
<div>

### 従来
- 計算量: `O(n^2)`
- メモリ: 多い
</div>
<div>

### 提案
- 計算量: `O(n \log n)`
- メモリ: 少ない
</div>
</div>
<div class="bottom">
<div class="conclusion"><strong>結論:</strong> 大規模データでも実用的な速度を実現。</div>
</div>

---

<!-- _class: kpi -->
# 実験結果
<div class="kpi-container">
<div><span class="kpi-value">97%</span><span class="kpi-label">精度</span></div>
<div><span class="kpi-value">10x</span><span class="kpi-label">高速化</span></div>
<div><span class="kpi-value">50%</span><span class="kpi-label">省メモリ</span></div>
</div>

---

<!-- _class: takeaway -->
# Takeaway
<div class="ta-main">計算量の改善が大規模データへの道をひらく</div>
<div class="ta-points">
<ul>
<li>計算量の改善により大規模データに対応</li>
<li>精度は従来と同等</li>
<li>OSS として公開予定</li>
</ul>
</div>

---

<!-- _class: references -->
# 参考文献
<ol>
<li><span class="author">Smith et al.</span> <span class="title">Fast Methods.</span> <span class="venue">NeurIPS 2024.</span></li>
<li><span class="author">山田</span> <span class="title">大規模データ処理入門.</span> <span class="venue">Ohmsha, 2025.</span></li>
</ol>

---

<!-- _class: end -->
# Thank You
Questions?
