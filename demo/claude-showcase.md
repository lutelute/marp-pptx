---
marp: true
---

<!-- _class: title -->
# 判断の速さは、設計できる
## PVホスティングキャパシティ評価の高速化
重信 颯人 — 福井大学 工学研究科 / 2026

---

<!-- _class: agenda -->
# 本日の 3 つの問い
<div class="agenda-list">
1. なぜ「当日回答」できないのか — 94 分の解剖
2. どう解いたか — 感度行列 LP × AC 潮流の分業を導く
3. どこまで言えるか — 実測・比較・限界
</div>

---

<!-- _class: divider -->
# 1. なぜ「当日回答」できないのか

---

<!-- _class: sections -->

# 94 分の正体は「総当たり × 二分探索」

## 課題の解剖｜計算量は 3 つの掛け算で決まる

<div class="sec">
  <span class="sec-title">構造｜ノード数 × 二分探索 × AC 潮流</span>
  <span class="sec-body">受入可否は「ノードごとに連系量を二分探索し、その都度 AC 潮流を解く」。1,000 ノード級では潮流計算が ==数万回== 走り、1 回数秒でも総計 94 分に達する。</span>
</div>

<div class="sec">
  <span class="sec-title">要件｜誤差 数%・応答 数分・既存 DB 継続</span>
  <span class="sec-body">実務が求めるのは、判定誤差 AC 比 **数 % 以内**・1 フィーダ **数分以内**・既存の系統データベースを **そのまま** 使えること。3 つを同時に満たす必要がある。</span>
</div>

<div class="sec">
  <span class="sec-title">既存代替の限界｜速い手法は精度か追従を欠く</span>
  <span class="sec-body">LinDistFlow 線形化は誤差 ±5.8%、ML 代理は系統変更で **再学習が必要**、モンテカルロは再サンプリング必須。「速さ・精度・追従」の3点を同時に満たす手法が無かった。</span>
</div>

---

<!-- _class: divider -->
# 2. どう解いたか

---

<!-- _class: sections -->
<!-- build -->

# 解法は 3 つの観察から導かれる

## 導出｜「どこを近似し、どこを正確に残すか」

<div class="sec">
  <span class="sec-title">観察 ①｜電圧感度はヤコビアンに既に入っている</span>
  <span class="sec-body">潮流計算は毎回ヤコビアンを作って捨てている。ここから感度行列 $S = \partial V / \partial p$ を抽出すれば、**追加の潮流計算なしで** 電圧応答の一次近似が手に入る。</span>
</div>

<div class="sec">
  <span class="sec-title">観察 ②｜二分探索は LP に置き換えられる</span>
  <span class="sec-body">感度が線形なら「電圧上限内で連系量最大化」は LP そのもの。ノードごとの二分探索（数十回の潮流）が ==1 回の一括最適化== に潰れる。</span>
</div>

<div class="sec">
  <span class="sec-title">観察 ③｜誤差は解の近傍でだけ生じる</span>
  <span class="sec-body">線形化誤差が効くのは制約に張り付く運転点の近傍のみ。**解の周辺だけ** AC 潮流で検証・補正すれば、精度は AC 級に戻る — これが分業の核心。</span>
</div>

---

<!-- _class: flow -->
<!-- build -->

# 3 つの観察が 1 つのループになる

## 仕組み｜LP が候補を出し、AC が正しさを保証する

```mermaid
flowchart LR
  S[感度行列 S<br>ヤコビアンから抽出] --> LP[線形最適化 LP<br>全ノード一括・1 秒未満]:::accent
  LP -->|候補解 p*| AC[AC 潮流計算<br>真の電圧・電流で検証]
  AC -->|違反なし| OUT([受入可否を当日回答]):::primary
  AC -.->|違反あり・再線形化（平均 3 往復）| LP
```

---

<!-- _class: equation -->
# 中核はただの線形計画になる
<div class="eq-main">

$$\max_{p \ge 0} \sum_{i} p_i \quad \text{s.t.} \quad V_{\min} \le V^{(0)} + S\,p \le V_{\max}$$

</div>
<div class="eq-desc">
<span>$p_i$</span><span>ノード $i$ の PV 連系量（決定変数）</span>
<span>$S$</span><span>電圧感度行列 — 基準潮流のヤコビアンから追加計算なしで取得</span>
<span>$V^{(0)}$</span><span>基準運転点の電圧 — 既存 DB の値をそのまま使う</span>
</div>

---

<!-- _class: divider -->
# 3. どこまで言えるか

---

<!-- _class: chart -->
<!-- _chart: column -->
<!-- source: 同一計算機（Xeon w5-2455X, 単スレッド）での実測 -->
# 計算時間はノード数にほぼ線形
## 実測｜300〜1,200 ノードの 3 規模で確認
| 規模 | 総当たり AC (分) | 提案手法 (分) |
|---|---|---|
| 300 | 12.1 | 0.9 |
| 600 | 28.4 | 1.3 |
| 1200 | 94.1 | 2.0 |

---

<!-- _class: kpi -->
# 速度・精度・収束性を同時に満たす
## 総合評価｜4 指標すべてが実務要件の内側
<div class="kpi-container">
<div><span class="kpi-value">47×</span><span class="kpi-label">高速化（94.1→2.0 分）</span></div>
<div><span class="kpi-value">±1.8%</span><span class="kpi-label">受入量誤差（AC 比）</span></div>
<div><span class="kpi-value">3 回</span><span class="kpi-label">平均 AC 補正反復</span></div>
</div>

---

<!-- _class: table-slide -->
# 速さ・精度・追従の 3 点を満たすのは提案のみ

比較（○=満たす／△=部分的／×=満たさない）

| 手法 | 速さ | 精度 | 系統変更追従 |
|---|---|---|---|
| 総当たり AC（基準） | × | ○ | ○ |
| LinDistFlow 線形化 | ○ | × | ○ |
| ML 代理モデル | ○ | △ | × |
| **提案手法** | ○ | ○ | ○ |

<div class="conclusion">
位置づけ: 純線形化の速度と AC 基準の精度の中間を取り、運用実務の「当日回答」要件を満たす唯一の構成。感度は再線形化のみで追従するため学習コストが無い。
</div>

---

<!-- _class: sections -->

# 言えること・言えないこと

## 限界の明示｜主張の範囲を先に切る

<div class="sec">
  <span class="sec-title">言えること｜放射状配電系統での当日回答</span>
  <span class="sec-body">6.6 kV 放射状フィーダ（〜1,200 ノード）で誤差 ±1.8%・2 分。申請が集中しても翌日持ち越しが消える。</span>
</div>

<div class="sec">
  <span class="sec-title">まだ言えないこと｜メッシュ系統・動的制約</span>
  <span class="sec-body">ループを含む系統や、タップ切替・逆潮流制御の動特性は検証外。感度行列の再抽出頻度も系統ごとの調整が要る。</span>
</div>

---

<!-- _class: takeaway -->
# まとめ
## 速さは感度行列で、正しさは AC 潮流で買う — 分業が「当日回答」を標準にする
- 根拠1: 1,200 ノード実測で 94.1 分 → 2.0 分（47 倍）、誤差 ±1.8%
- 根拠2: 収束は平均 3 反復 — 分業の仮定（誤差は解の近傍のみ）が実測で成立
- 限界: 放射状系統に限る。メッシュ・動的制約は次の課題

---

<!-- _class: end -->
# ご清聴ありがとうございました
質問・議論歓迎 — 実装は marp-pptx デモとして公開
