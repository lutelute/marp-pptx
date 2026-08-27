---
marp: true
theme: academic
paginate: true
---

<!-- _class: graphical-abstract -->

# 本研究の全体像

## 一枚で｜課題 → 提案 → 成果

<div class="ga-problem">
  <span class="ga-label">課題</span>
  <span class="ga-body">PV 受入可否の判定は総当たり AC × 二分探索で **94 分**。申請ペースに計算が追いつかず、回答は翌日持ち越しになっている。</span>
</div>

<div class="ga-method">
  <span class="ga-label">提案</span>
  <span class="ga-steps">感度行列 → LP 一括 → AC 検証</span>
  <span class="ga-body">ヤコビアンから抽出した感度行列で LP が候補を一括生成し、AC 潮流は違反時のみ再検証。速さと正しさを ==分業== する。</span>
</div>

<div class="ga-result">
  <span class="ga-label">成果</span>
  <span class="ga-kpi">47×</span>
  <span class="ga-body">94.1 分 → 2.0 分（1,200 ノード実測）
誤差 ±1.8%・平均 3 反復で収束</span>
</div>

<div class="ga-foot">実測条件: Xeon w5-2455X 単スレッド／6.6 kV 放射状フィーダ 300〜1,200 ノード</div>
