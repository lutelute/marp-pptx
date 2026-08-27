---
marp: true
theme: academic
paginate: true
---

<!-- _class: flow -->

# 提案手法の全体構成

## 反復補正｜LP が候補を出し、AC 潮流が正しさを保証する

```mermaid
flowchart LR
  S[感度行列 S<br>ヤコビアンから構築] --> LP[線形最適化 LP<br>全ノード一括]:::accent
  LP -->|候補解 p*| AC[AC 潮流計算<br>電圧・電流を検証]
  AC -->|違反なし| OUT([受入可否を回答]):::primary
  AC -.->|違反あり・再線形化| LP
```
