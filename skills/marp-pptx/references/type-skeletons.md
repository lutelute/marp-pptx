# 型ごとの HTML 骨組みリファレンス

marp-pptx の全 52 型の正確な構造。`<!-- _class: -->` とその下の HTML を**崩さずに**埋める。
（このファイルは `scripts/gen-skeletons.py` が `src/marp_pptx/data/templates/` から自動生成）

## 構造（structure）

### `cols-2` — 並列・対比
<!-- □ □ · 2つを同じ重みで比較するとき -->

```markdown
<!-- _class: cols-2 -->

# 2カラムレイアウト

<div class="columns">
<div>

## 左カラム

- 項目1: 説明テキスト
- 項目2: 説明テキスト
- 項目3: 説明テキスト

$$E = mc^2$$

</div>
<div>

## 右カラム

<div class="box">

**ポイント**: ここにハイライトしたい内容を記述する。`box` クラスで囲む。

</div>

<div class="box-accent">

**注意**: `box-accent` でアクセントカラー付きのボックスになる。

</div>

</div>
</div>
```

### `cols-3` — 分類・カテゴリ
<!-- □ □ □ · 3つの側面を示すとき -->

```markdown
<!-- _class: cols-3 -->

# 3カラムレイアウト

<div class="columns">
<div>

## 手法A

<div class="box-primary">

- 特徴1
- 特徴2
- 精度: 92.3%

</div>

</div>
<div>

## 手法B

<div class="box-primary">

- 特徴1
- 特徴2
- 精度: 94.1%

</div>

</div>
<div>

## 手法C (提案)

<div class="box-accent">

- 特徴1
- 特徴2
- 精度: **96.7%**

</div>

</div>
</div>

<div class="footnote">※ 精度は検証データセットでの結果</div>
```

### `sandwich` — 概要→詳細→結論
<!-- ─ □□ ─ · フレーミングが必要なとき -->

```markdown
<!-- _class: sandwich -->

# サンドイッチレイアウト（リード文＋マルチカラム＋結論文）

<div class="top">

<p class="lead">リード文エリア。上部全幅でスライドの文脈や背景を記述する。問題設定や前提条件の導入に最適。</p>

</div>

<div class="columns c2">
<div>

## アプローチ1
- 手順A → 手順B → 手順C
- 計算量: $O(n \log n)$

</div>
<div>

## アプローチ2
- 手順X → 手順Y
- 計算量: $O(n^2)$

</div>
</div>

<div class="bottom">

<div class="conclusion">

**結論**: アプローチ1は計算量で優位であり、大規模データに対するスケーラビリティが高い。次セクションで実験的に検証する。

</div>

</div>
```

### `sandwich-3col` — 概要→3並列→結論
<!-- ─ □□□ ─ · 共通設定＋3条件＋考察を1枚にまとめるとき -->

```markdown
<!-- _class: sandwich -->

# 3カラム・サンドイッチ

<div class="top">

上部全幅エリア。実験条件の共通設定やデータセットの説明など。

</div>

<div class="columns c3">
<div>

### データセットA
- サンプル数: 10,000
- 精度: 95.2%

</div>
<div>

### データセットB
- サンプル数: 50,000
- 精度: 96.8%

</div>
<div>

### データセットC
- サンプル数: 100,000
- 精度: 97.1%

</div>
</div>

<div class="bottom">

**考察**: データ量の増加に伴い精度は向上するが、50K→100K では改善幅が縮小。

</div>
```

### `split-text` — 二面性・補完
<!-- □│□ · 左右で補完的な内容を示すとき -->

```markdown
<!-- _class: split-text -->

# 仮説と結果

<div class="sp-left">
  <span class="sp-label">仮説</span>
  <span class="sp-body">動的スパースマスクを導入することで、計算量を削減しつつ精度を維持できる。マスクパターンはタスクに応じて自動的に最適化される。</span>
</div>

<div class="sp-right">
  <span class="sp-label">結果</span>
  <span class="sp-body">3つのベンチマークで精度を維持しつつ、計算量を平均 60% 削減。マスクの可視化により、タスク依存のパターンを確認。</span>
</div>
```

### `card-grid` — 均質な要素の一覧
<!-- □□ □□ · 同種のアイテムを並べるとき -->

```markdown
<!-- _class: card-grid -->

# 実験条件

<div class="cg-container">

<div class="cg-card">
  <span class="cg-title">条件 A: 短系列</span>
  <span class="cg-body">系列長 512。標準的な NLU タスク。BERT-base と同等設定。</span>
</div>

<div class="cg-card">
  <span class="cg-title">条件 B: 中系列</span>
  <span class="cg-body">系列長 2048。文書分類・要約タスク。GPT-2 と同等設定。</span>
</div>

<div class="cg-card">
  <span class="cg-title">条件 C: 長系列</span>
  <span class="cg-body">系列長 8192。長距離依存タスク。Longformer と同等設定。</span>
</div>

<div class="cg-card">
  <span class="cg-title">条件 D: 超長系列</span>
  <span class="cg-body">系列長 16384。書籍レベル。本研究の主要ターゲット。</span>
</div>

</div>
```

### `split-panel` — ハーフブリード色面パネル
<!-- ▮左色面│本文 · 画面端まで塗った色面に主張を白抜きし、右に本文を置くとき（キーノート級の1枚） -->

```markdown
<!-- _class: split-panel -->

# 速さと正しさは、もう交換条件ではない

## PROPOSAL

<div class="sp-body">
- 感度行列 LP が **1 秒未満** で候補を出す
- AC 潮流は違反時のみ — 平均 ==3 往復== で収束
- 判定誤差は AC 比 ±1.8%、当日回答が標準になる
- 既存の系統データベースはそのまま使える
</div>
```

### `graphical-abstract` — グラフィカルアブストラクト
<!-- □→□→▪ 一枚絵 · 表紙直後に課題→手法→成果を1枚の図で示すとき（研究発表の定番） -->

```markdown
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
```

### `figure-full` — 論文図の全面表示
<!-- ▣ 最大画角の図 · 論文の特徴的な図を余白0.25inまで最大サイズで見せるとき（図が主役の1枚） -->

```markdown
<!-- _class: figure-full -->
<!-- source: Vaswani et al. (2017), Fig. 1 -->

# Transformer の全体アーキテクチャ

![w:1200](figures/transformer-architecture.png)

<div class="caption">エンコーダ・デコーダとも自己注意＋FFN の積層のみで構成される</div>
```

## 時間（temporal）

### `timeline-h` — 時系列・経過
<!-- ●─●─●─● · 時間の流れを示すとき -->

```markdown
<!-- _class: timeline-h -->

# 研究の歴史的流れ（横型タイムライン）

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
    <span class="tl-h-year">2015</span>
    <span class="tl-h-text">ResNet</span>
    <div class="tl-h-detail">残差接続<br>152層の深層化</div>
  </div>
</div>

<div class="tl-h-item">
  <div class="tl-h-block">
    <span class="tl-h-year">2017</span>
    <span class="tl-h-text">Transformer</span>
    <div class="tl-h-detail">Self-attention<br>並列処理革命</div>
  </div>
</div>

<div class="tl-h-item">
  <div class="tl-h-block">
    <span class="tl-h-year">2019</span>
    <span class="tl-h-text">BERT</span>
    <div class="tl-h-detail">事前学習<br>双方向 Transformer</div>
  </div>
</div>

<div class="tl-h-item highlight">
  <div class="tl-h-block">
    <span class="tl-h-year">2024</span>
    <span class="tl-h-text bold">本研究</span>
    <div class="tl-h-detail">Sparse + 収束保証<br>計算量60%削減</div>
  </div>
</div>

</div>
```

### `timeline-v` — 手順・段階
<!-- ●│●│● · 縦の時系列を示すとき -->

```markdown
<!-- _class: timeline -->

# 研究の歴史的流れ（縦型タイムライン）

<div class="tl-container">

<div class="tl-item">
  <span class="tl-year">2012</span>
  <span class="tl-text">AlexNet — 深層学習による画像認識の躍進</span>
  <div class="tl-detail">ImageNet 大規模データでCNNが従来手法を大幅に上回る</div>
</div>

<div class="tl-item">
  <span class="tl-year">2014</span>
  <span class="tl-text">GAN — 生成モデルの新パラダイム</span>
  <div class="tl-detail">Generator/Discriminator の敵対的学習フレームワーク</div>
</div>

<div class="tl-item">
  <span class="tl-year">2015</span>
  <span class="tl-text">ResNet — 残差接続による超深層ネットワーク</span>
  <div class="tl-detail">152層の深層化を実現、勾配消失問題を解決</div>
</div>

<div class="tl-item">
  <span class="tl-year">2017</span>
  <span class="tl-text">Transformer — 自己注意機構の提案</span>
  <div class="tl-detail">RNN不要の並列処理アーキテクチャ、NLP分野を革新</div>
</div>

<div class="tl-item highlight">
  <span class="tl-year">2024</span>
  <span class="tl-text bold">本研究 — Transformer の効率的拡張</span>
  <div class="tl-detail">Sparse Attention + 収束保証で計算量60%削減</div>
</div>

</div>
```

### `steps` — プロセス・手順
<!-- ①→②→③ · 順にやる手順を示すとき -->

```markdown
<!-- _class: steps -->

# 感情ベクトルの抽出手順

<div class="st-container">

<div class="st-step">
  <span class="st-num">1</span>
  <span class="st-title">感情語の選定</span>
  <span class="st-body">171種の感情語を選定し、各感情に対応するプロンプトを設計</span>
</div>

<div class="st-step">
  <span class="st-num">2</span>
  <span class="st-title">ストーリー生成</span>
  <span class="st-body">各感情について1200本のストーリーを LLM で生成</span>
</div>

<div class="st-step">
  <span class="st-num">3</span>
  <span class="st-title">残差ストリーム取得</span>
  <span class="st-body">生成過程の中間層から残差ストリームベクトルを抽出</span>
</div>

<div class="st-step">
  <span class="st-num">4</span>
  <span class="st-title">ノイズ除去</span>
  <span class="st-body">平均ベクトルと交絡因子の影響を project out して精製</span>
</div>

</div>
```

### `before-after` — 変化・改善
<!-- □ → □ · ビフォーアフターを示すとき -->

```markdown
<!-- _class: before-after -->

# 改善効果

<div class="ba-before">
  <span class="ba-label">Before</span>
  <span class="ba-body">Full Attention で全ペアを計算。系列長 4096 で OOM が頻発。学習に 72 時間。</span>
</div>

<div class="ba-after">
  <span class="ba-label">After</span>
  <span class="ba-body">Sparse Attention で計算量 60% 削減。系列長 16384 でも安定動作。学習 30 時間。</span>
</div>
```

### `history` — 沿革・文脈
<!-- 年│出来事 · 歴史的経緯を示すとき -->

```markdown
<!-- _class: history -->

# 研究の沿革

<div class="hs-container">

<div class="hs-item">
  <span class="hs-year">2017</span>
  <span class="hs-event">Transformer アーキテクチャの提案 (Vaswani et al.)</span>
</div>

<div class="hs-item">
  <span class="hs-year">2019</span>
  <span class="hs-event">BERT による事前学習モデルの普及</span>
</div>

<div class="hs-item">
  <span class="hs-year">2022</span>
  <span class="hs-event">Flash Attention による IO 最適化</span>
</div>

<div class="hs-item">
  <span class="hs-year">2025</span>
  <span class="hs-event">スパース注意機構の理論的保証（本研究）</span>
</div>

</div>
```

## 収束・拡散（convergence）

### `funnel` — 絞り込み・選別
<!-- ▽ · 多→少の過程を見せるとき -->

```markdown
<!-- _class: funnel -->

# データ処理パイプライン

<div class="fn-container">

<div class="fn-stage">
  <span class="fn-label">Raw Data</span>
  <span class="fn-value">1,000,000 件</span>
</div>

<div class="fn-stage">
  <span class="fn-label">前処理済み</span>
  <span class="fn-value">850,000 件</span>
</div>

<div class="fn-stage">
  <span class="fn-label">品質フィルタ後</span>
  <span class="fn-value">500,000 件</span>
</div>

<div class="fn-stage">
  <span class="fn-label">最終データセット</span>
  <span class="fn-value">100,000 件</span>
</div>

</div>
```

### `stack` — 積み上げ・累積
<!-- □ □ □ 積層 · レイヤー構造を示すとき -->

```markdown
<!-- _class: stack -->

# システムアーキテクチャ

<div class="sk-container">

<div class="sk-layer">
  <span class="sk-name">Application Layer</span>
  <span class="sk-desc">推論 API、バッチ処理、リアルタイム予測</span>
</div>

<div class="sk-layer">
  <span class="sk-name">Model Layer</span>
  <span class="sk-desc">Sparse Attention Block、マスク生成器、Embedding</span>
</div>

<div class="sk-layer">
  <span class="sk-name">Framework Layer</span>
  <span class="sk-desc">PyTorch 2.0、CUDA カーネル、Mixed Precision</span>
</div>

<div class="sk-layer">
  <span class="sk-name">Infrastructure</span>
  <span class="sk-desc">NVIDIA A100 x 4、NVLink、高速ストレージ</span>
</div>

</div>
```

### `overview` — 全体像と部分
<!-- 大□＋小□群 · 全体→詳細の構造を示すとき -->

```markdown
<!-- _class: overview -->

# 概要と主要な知見

<div class="ov-lead">ここに研究の概要を1〜2文で記述する。モデル内部の感情ベクトルを抽出し、その挙動を分析した結果、感情操作がモデルの行動に直接影響を与えることが示された。</div>

![w:700](assets/architecture.svg)

<div class="caption">Fig. 1. 全体の構成図。ここに図の読み方を簡潔に。</div>

<div class="ov-points">
<li>知見 1: 171種の感情ベクトルを抽出し、直感的なクラスタを形成</li>
<li>知見 2: 入力に応じて感情の内部表現が動的に変化</li>
<li>知見 3: 感情の強制操作により望ましくない行動が増減</li>
</div>
```

### `highlight` — 強調・焦点
<!-- ███ ■ ███ · 1つだけ際立たせるとき -->

```markdown
<!-- _class: highlight -->

# 核心メッセージ

<div class="hl-text">動的スパースマスクにより、計算量を 60% 削減しながら精度を維持できる</div>
```

## 評価・判断（evaluation）

### `pros-cons` — 賛否・長短
<!-- ＋│− · 判断材料を示すとき -->

```markdown
<!-- _class: pros-cons -->

# 提案手法の利点と制約

<div class="pc-pros">
<li>計算量を 60% 削減</li>
<li>既存パイプラインへのドロップイン置換が可能</li>
<li>理論的な収束保証あり</li>
<li>GPU 並列化に対応</li>
</div>

<div class="pc-cons">
<li>短系列タスクではオーバーヘッドが相対的に大きい</li>
<li>マスク生成器の追加学習が必要</li>
<li>動的マスクの解釈性が限定的</li>
</div>
```

### `zone-compare` — 比較評価
<!-- □ vs □ · 2つを評価比較するとき -->

```markdown
<!-- _class: zone-compare -->

# 手法の比較

<div class="zc-container">

<div class="zc-left">
  <span class="zc-label">従来手法</span>
  <span class="zc-body">バッチ処理ベース。計算量 $O(n^2)$。精度は中程度だが安定性が高い。大規模データではスケーラビリティに課題。</span>
</div>

<div class="zc-vs">VS</div>

<div class="zc-right">
  <span class="zc-label">提案手法</span>
  <span class="zc-body">ストリーム処理対応。計算量 $O(n \log n)$。精度向上しつつリアルタイム処理を実現。GPU 並列化に対応。</span>
</div>

</div>
```

### `zone-matrix` — 二軸評価
<!-- □□ □□ (2x2) · 2軸で分類するとき -->

```markdown
<!-- _class: zone-matrix -->

# 研究アプローチの分類

<div class="zm-container">

<div class="zm-ylabel">精度</div>
<div class="zm-xlabel">計算コスト</div>

<div class="zm-cell zm-tl">
  <span class="zm-label">理想領域</span>
  <span class="zm-body">高精度 + 低コスト。提案手法が目指すターゲット。</span>
</div>

<div class="zm-cell zm-tr">
  <span class="zm-label">力技</span>
  <span class="zm-body">高精度だが高コスト。大規模モデルや総当たり探索が該当。</span>
</div>

<div class="zm-cell zm-bl">
  <span class="zm-label">ベースライン</span>
  <span class="zm-body">低コスト・低精度。ルールベースや単純なヒューリスティクス。</span>
</div>

<div class="zm-cell zm-br">
  <span class="zm-label">非効率</span>
  <span class="zm-body">高コストなのに低精度。設計上の問題があるアプローチ。</span>
</div>

</div>
```

### `kpi` — 定量評価
<!-- 数字 数字 数字 · KPI・数値を強調するとき -->

```markdown
<!-- _class: kpi -->

# 主要指標

<div class="kpi-container">

<div class="kpi-item">
  <span class="kpi-value">89.4%</span>
  <span class="kpi-label">分類精度</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">2.4x</span>
  <span class="kpi-label">処理速度向上</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">-40%</span>
  <span class="kpi-label">メモリ使用量削減</span>
</div>

<div class="kpi-item">
  <span class="kpi-value">127M</span>
  <span class="kpi-label">パラメータ数</span>
</div>

</div>
```

### `result` — 成果報告
<!-- 結論＋根拠 · 実験結果を報告するとき -->

```markdown
<!-- _class: result -->

# ステアリング結果：感情操作の影響

<div class="rs-lead">感情ベクトルのステアリング強度ごとに望ましくない行動の割合を観察した。desperate を正方向にステアリングするとブラックメール行為が大幅に増加する。</div>

<div class="rs-figure">

![w:500](assets/learning-curve.svg)

<div class="caption">Fig. 2. ステアリング強度と行動変化の関係</div>

</div>

<div class="rs-analysis">
<li>desperate の正方向ステアリングでブラックメール行為が増加</li>
<li>calm の正方向ステアリングで同行為が劇的に減少</li>
<li>ステアリング強度と行動変化は単調な関係ではない</li>
<li>中間強度で最も顕著な効果が観察される</li>
</div>
```

### `result-dual` — 二つの成果を並列
<!-- 結果□│結果□ · 2つの結果を比較するとき -->

```markdown
<!-- _class: result-dual -->

# 実験結果の比較

<div class="results">

<div class="result-item">

![w:500](assets/learning_curve.png)

<div class="caption"><span class="fig-num">Fig. 2.</span> 学習曲線の比較</div>

</div>

<div class="result-item">

![w:500](assets/sparse_pattern.png)

<div class="caption"><span class="fig-num">Fig. 3.</span> スパースパターンの可視化</div>

</div>

</div>
```

### `multi-result` — 複数成果の一覧
<!-- 結果□□□ · 複数の結果を一覧するとき -->

```markdown
<!-- _class: multi-result -->

# 定量評価

<div class="mr-container">

<div class="mr-item">
  <span class="mr-metric">Perplexity</span>
  <span class="mr-value">17.9</span>
  <span class="mr-desc">Full Attention (18.3) と同等の言語モデリング性能</span>
</div>

<div class="mr-item">
  <span class="mr-metric">Speedup</span>
  <span class="mr-value">2.4x</span>
  <span class="mr-desc">系列長 8192 での推論速度。A100 単体での測定</span>
</div>

<div class="mr-item">
  <span class="mr-metric">Memory</span>
  <span class="mr-value">-40%</span>
  <span class="mr-desc">ピークメモリ使用量の削減率。バッチサイズ 64</span>
</div>

</div>
```

### `big-number` — 単一指標の強調
<!-- 巨大数字 · ひとつの数字を主役にするとき -->

```markdown
<!-- _class: big-number -->
<!-- source: 社内ベンチマーク 2026 (n=1000) -->

# 主要成果

<div class="big-number">
  <span class="bn-value">89.4%</span>
  <span class="bn-label">分類精度</span>
  <span class="bn-caption">従来手法から +4.2pt 改善</span>
</div>
```

### `chart` — データの可視化
<!-- 棒/折れ線グラフ · 数値データをグラフで見せるとき(編集可) -->

```markdown
<!-- _class: chart -->
<!-- _chart: column -->

# 系列長ごとの計算時間（相対）

| 系列長 | Transformer | Ours |
|---|---|---|
| 1024 | 1.0 | 0.6 |
| 4096 | 4.2 | 1.1 |
| 16384 | 18.5 | 2.3 |

<div class="chart-caption">同一ハードウェアで測定。Ours は Okabe-Ito 安全色で強調。</div>
```

## 知識・定義（knowledge）

### `definition` — 概念の定義
<!-- 用語：説明 · 用語を定義するとき -->

```markdown
<!-- _class: definition -->

# 用語定義

<div class="df-term">スパースアテンション (Sparse Attention)</div>

<div class="df-body">入力系列の全ペア間ではなく、選択的に注意重みを計算する手法。計算量を $O(n^2)$ から $O(n\sqrt{n})$ 以下に削減しつつ、タスク性能を維持することを目指す。</div>

<div class="df-note">関連概念: Self-Attention, Multi-Head Attention, Linear Attention</div>
```

### `equation` — 数理的真理
<!-- $式$ 中央 · 数式が主役のとき -->

```markdown
<!-- _class: equation -->

# 数式スライド — 基本

<div class="eq-main">

$$\mathcal{L}(\theta) = -\frac{1}{N}\sum_{i=1}^{N} \left[ y_i \log \hat{y}_i + (1-y_i)\log(1-\hat{y}_i) \right]$$

</div>

<div class="eq-desc">
  <span class="sym">$\mathcal{L}(\theta)$</span>
  <span>損失関数（交差エントロピー）</span>
  <span class="sym">$\theta$</span>
  <span>モデルパラメータ</span>
  <span class="sym">$N$</span>
  <span>サンプル数</span>
  <span class="sym">$y_i$</span>
  <span>正解ラベル（0 or 1）</span>
  <span class="sym">$\hat{y}_i$</span>
  <span>モデルの予測確率</span>
</div>
```

### `equations` — 式の体系
<!-- $式$ $式$ $式$ · 連立方程式・最適化問題のとき -->

```markdown
<!-- _class: equations -->

# 最適化問題の定式化

<div class="eq-system">
  <div class="row">
    <span class="label">minimize</span>

$$f(x) = \tfrac{1}{2} x^\top Q x + c^\top x$$

  </div>
  <div class="row">
    <span class="label">subject to</span>

$$A x \le b$$

  </div>
  <div class="row">
    <span class="label"></span>

$$C x = d$$

  </div>
  <div class="row">
    <span class="label"></span>

$$x \ge 0$$

  </div>
</div>

<div class="eq-desc">
  <span class="sym">$x \in \mathbb{R}^{n}$</span>
  <span>決定変数</span>
  <span class="sym">$Q \in \mathbb{S}^{n}_{+}$</span>
  <span>半正定値行列（二次コスト）</span>
  <span class="sym">$c \in \mathbb{R}^{n}$</span>
  <span>線形コスト係数</span>
  <span class="sym">$A \in \mathbb{R}^{m \times n},\ b \in \mathbb{R}^{m}$</span>
  <span>不等式制約</span>
  <span class="sym">$C \in \mathbb{R}^{p \times n},\ d \in \mathbb{R}^{p}$</span>
  <span>等式制約</span>
</div>

<div class="footnote">$Q \succeq 0$ ならば凸二次計画問題（QP）として扱える。</div>
```

### `equation-annotated` — 数式＋記号注釈
<!-- $式$ ＋ 凡例 · 数式の各記号を注釈付きで解説するとき -->

```markdown
<!-- _class: equation -->

# 数式スライド — アノテーション付き

<div class="eq-main">

$$\hat{y} = \underbrace{\sigma}_{\text{活性化}} \left( \overbrace{W}^{\text{重み}} \cdot \underbrace{x}_{\text{入力}} + \overbrace{b}^{\text{バイアス}} \right)$$

</div>

<div class="eq-desc">
  <span class="sym">$\hat{y}$</span>
  <span>出力（予測値）</span>
  <span class="sym">$\sigma(\cdot)$</span>
  <span>シグモイド活性化関数: $\sigma(z) = \frac{1}{1+e^{-z}}$</span>
  <span class="sym">$W$</span>
  <span>重み行列 $\in \mathbb{R}^{m \times n}$</span>
  <span class="sym">$x$</span>
  <span>入力ベクトル $\in \mathbb{R}^{n}$</span>
  <span class="sym">$b$</span>
  <span>バイアスベクトル $\in \mathbb{R}^{m}$</span>
</div>

<div class="footnote">KaTeX の \underbrace / \overbrace を使って数式内に直接アノテーションを記述</div>
```

### `equation-highlight` — 数式の着目部を強調
<!-- $式$ 領域強調 · 数式の特定項を色で強調して読み解くとき -->

```markdown
<!-- _class: equation -->

# 数式スライド — 領域ハイライト

<div class="eq-main">

$$\text{Attention}(Q, K, V) = \text{softmax}\!\left(\frac{\colorbox{#fff3cd}{$QK^\top$}}{\colorbox{#cce5ff}{$\sqrt{d_k}$}}\right)\colorbox{#fff3cd}{$V$}$$

</div>

<div class="eq-desc">
  <span class="sym"><span class="eq-highlight">$QK^\top$</span></span>
  <span>クエリとキーの内積 → 類似度スコア</span>
  <span class="sym"><span class="eq-highlight-b">$\sqrt{d_k}$</span></span>
  <span>スケーリング係数（勾配消失を防止）</span>
  <span class="sym"><span class="eq-highlight">$V$</span></span>
  <span>バリュー行列 → 最終出力の重み付け対象</span>
  <span class="sym">$\text{softmax}$</span>
  <span>正規化 → 注意重みの確率分布化</span>
</div>

<div class="footnote">Vaswani et al., "Attention Is All You Need," NeurIPS 2017</div>
```

### `diagram` — 構造の可視化
<!-- 図＋説明 · 図で構造を説明するとき -->

```markdown
<!-- _class: diagram -->

# アーキテクチャ図

![w:900](assets/architecture.svg)

<div class="caption">Fig. 1. 提案手法の全体構成。入力から出力までのデータフロー。</div>
```

### `annotation` — 詳細解説
<!-- 図＋注釈 · 図に注釈を付けるとき -->

```markdown
<!-- _class: annotation -->

# 図の注釈

<div class="an-figure">

![w:500](assets/architecture.svg)

</div>

<div class="an-notes">
<li>入力層で BPE トークナイズを適用</li>
<li>中間層でスパースマスクを動的生成</li>
<li>出力層で次トークン予測を実行</li>
<li>マスク生成器は 2M パラメータ</li>
</div>
```

### `code` — 実装・手続き
<!-- コードブロック · コードを見せるとき -->

```markdown
<!-- _class: code -->

# コード例

<div class="cd-code">

```python
class SparseAttention(nn.Module):
    def __init__(self, d_model, n_heads):
        super().__init__()
        self.mask_gen = MaskGenerator(d_model)
        self.attn = nn.MultiheadAttention(
            d_model, n_heads
        )

    def forward(self, x):
        mask = self.mask_gen(x)
        return self.attn(x, x, x,
                         attn_mask=mask)
```

</div>

<div class="cd-desc">MaskGenerator がスパースマスクを動的生成し、標準の MultiheadAttention に渡す。既存コードへの変更は最小限。</div>
```

### `blocks` — 定理・定義ブロック
<!-- ▬題│本文 ×N · beamer 流の定理環境（theorem/example/alert）を並べるとき -->

```markdown
<!-- _class: blocks -->

# 定理ブロック（beamer 流）

<div class="bk-container">

<div class="bk theorem">
  <span class="bk-title">定理 1（収束性）</span>
  <span class="bk-body">提案する反復法は任意の初期点から線形レートで収束する。すなわち $\|x_{k+1} - x^*\| \le \rho \|x_k - x^*\|$（$0 < \rho < 1$）。</span>
</div>

<div class="bk example">
  <span class="bk-title">例（二次関数）</span>
  <span class="bk-body">$f(x) = x^2$ では $\rho = 1/2$ となり、10 反復で誤差はおよそ $10^{-3}$ 倍に縮む。</span>
</div>

<div class="bk alert">
  <span class="bk-title">注意</span>
  <span class="bk-body">ステップ幅が $2/L$ を超えると発散する（$L$ は勾配のリプシッツ定数）。</span>
</div>

</div>
```

### `sections` — 高密度トピック積層
<!-- ■リード│本文 ×N · 1枚に2〜4トピックを「色付きリード行＋本文」の帯で積むとき（公聴会流の高密度） -->

```markdown
<!-- _class: sections -->

# 需給運用における課題

<div class="sec">
  <span class="sec-title">ダックカーブ現象｜① 供給側の負担増加</span>
  <span class="sec-body">PVS の大量導入で **正味電力需要** がダックカーブ形状へ変化。正午の谷と夕方の急峻な立ち上がりに、発電機側の調整力だけでは追随できなくなりつつある。</span>
</div>

<div class="sec">
  <span class="sec-title">需給逼迫｜② 予備力の不足</span>
  <span class="sec-body">下げしろ・上げしろの両方向で予備力が逼迫。発電機出力の下限値と最大変化率が制約となり、==需要側の参加== が不可欠になる。</span>
</div>

<div class="sec">
  <span class="sec-title">要請｜③ 意思決定の定量的根拠</span>
  <span class="sec-body">技術選択・設備容量計画・運転スケジューリングの 3 つの意思決定に、評価指標（コスト・CO₂ 排出量）の定量値を与える枠組みが必要。</span>
</div>
```

### `paper` — 文献報告の書誌カード
<!-- ▤書誌+理由+要点 · 輪読・文献調査で紹介論文の書誌情報・選定理由・要点を1枚にするとき -->

```markdown
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
```

## 流れ・構造（flow）

### `flow` — ブロック図・ループ図
<!-- □→□→□ +loop · mermaid flowchart 記法から編集可能なブロック図/機器構成図/反復ループ図を描くとき -->

```markdown
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
```

### `zone-flow` — フロー・因果
<!-- □→□→□ · 原因→結果の流れを示すとき -->

```markdown
<!-- _class: zone-flow -->

# 処理パイプライン

<div class="zf-container">

<div class="zf-box">
  <span class="zf-label">Data Collection</span>
  <span class="zf-body">公開データセットの取得、ノイズ除去、フォーマット統一</span>
</div>

<div class="zf-box">
  <span class="zf-label">Feature Engineering</span>
  <span class="zf-body">次元削減、正規化、ドメイン特徴量の設計</span>
</div>

<div class="zf-box">
  <span class="zf-label">Model Training</span>
  <span class="zf-body">交差検証、ハイパーパラメータ探索、アンサンブル構築</span>
</div>

<div class="zf-box">
  <span class="zf-label">Evaluation</span>
  <span class="zf-body">テストセット評価、ベースライン比較、統計検定</span>
</div>

</div>
```

### `zone-process` — プロセス＋詳細
<!-- □→□→□ (詳細) · 詳細付きプロセスを示すとき -->

```markdown
<!-- _class: zone-process -->

# 実験手順

<div class="zp-container">

<div class="zp-step">
  <span class="zp-num">1</span>
  <span class="zp-title">データ収集</span>
  <span class="zp-body">公開データセットから10万件のサンプルを取得し、ノイズ除去を実施。</span>
</div>

<div class="zp-step">
  <span class="zp-num">2</span>
  <span class="zp-title">前処理</span>
  <span class="zp-body">正規化、欠損値補完、特徴量エンジニアリングを適用。</span>
</div>

<div class="zp-step">
  <span class="zp-num">3</span>
  <span class="zp-title">モデル学習</span>
  <span class="zp-body">5-fold交差検証でハイパーパラメータを最適化。学習率スケジューリング導入。</span>
</div>

<div class="zp-step">
  <span class="zp-num">4</span>
  <span class="zp-title">評価・比較</span>
  <span class="zp-body">テストセットでAUC、F1スコアを算出し、3つのベースラインと比較。</span>
</div>

</div>
```

### `agenda` — 予定・構成
<!-- 1. 2. 3. · 発表の構成を示すとき -->

```markdown
<!-- _class: agenda -->

# 発表の構成

<div class="agenda-list">

1. 研究背景と課題
2. 提案手法
3. 実験と結果
4. 考察
5. まとめと今後の課題

</div>
```

### `checklist` — 進捗・完了状態
<!-- ☑ ☑ ☐ · タスクの状態を示すとき -->

```markdown
<!-- _class: checklist -->

# 実験チェックリスト

<div class="cl-container">
<li class="done">データセットの前処理と分割</li>
<li class="done">ベースラインモデルの実装と検証</li>
<li class="done">提案手法の実装</li>
<li>ハイパーパラメータ探索</li>
<li>統計的有意性検定</li>
<li>アブレーション実験</li>
</div>
```

## ナラティブ（narrative）

### `quote` — 権威・声
<!-- 「　」 · 引用を示すとき -->

```markdown
<!-- _class: quote -->

# 引用スライド

<div class="qt-text">科学とは、知識の体系ではなく、思考の方法である。</div>

<div class="qt-source">Carl Sagan, The Demon-Haunted World (1995)</div>
```

### `profile` — 人物紹介
<!-- 写真＋経歴 · 人物を紹介するとき -->

```markdown
<!-- _class: profile -->

# 発表者紹介

<div class="pf-container">

<div class="pf-name">山田 太郎</div>

<div class="pf-affiliation">〇〇大学 工学研究科 情報工学専攻</div>

<div class="pf-bio">
<li>研究テーマ: 効率的な注意機構の設計と理論解析</li>
<li>所属: 自然言語処理研究室 (指導教員: 佐藤教授)</li>
<li>学会活動: ACL 2025 採択、NeurIPS 2024 Workshop</li>
<li>連絡先: yamada@example.ac.jp</li>
</div>

</div>
```

### `takeaway` — 持ち帰ってほしい1つ
<!-- ★ メッセージ · キーメッセージを伝えるとき -->

```markdown
<!-- _class: takeaway -->

# キーメッセージ

<div class="ta-main">動的スパース注意機構は、精度を犠牲にせず計算効率を大幅に改善できる</div>

<div class="ta-points">
<li>3つのベンチマークで SOTA と同等の精度を達成</li>
<li>計算量を平均 60% 削減し、長系列タスクへの適用を実現</li>
<li>既存の効率化手法 (Flash Attention) と直交的に併用可能</li>
</div>
```

### `statement` — 断言・転換
<!-- 全画面 1文 · 1メッセージを大きく言い切るとき(dark可) -->

```markdown
<!-- _class: statement -->

本研究は、スパース注意機構に初めて理論的保証を与える。
```

### `panorama` — インパクト・没入
<!-- 横幅画像 · 大きな画像で印象づけるとき -->

```markdown
<!-- _class: panorama -->

# システム概要

<div class="pn-text">左側にリードテキストを配置し、全体像を説明する。右側には概念図やアーキテクチャ図を大きく表示する。テキストと図の組み合わせで、直感的な理解を促す。</div>

![w:600](assets/architecture.svg)
```

### `gallery-img` — ビジュアル一覧
<!-- 画像群 · 複数画像を並べるとき -->

```markdown
<!-- _class: gallery-img -->

# 実験結果ギャラリー

<div class="gi-container">

<div class="gi-item">

![w:400](assets/architecture.svg)

<div class="gi-caption">条件 A の結果</div>
</div>

<div class="gi-item">

![w:400](assets/learning-curve.svg)

<div class="gi-caption">条件 B の結果</div>
</div>

<div class="gi-item">

![w:400](assets/architecture.svg)

<div class="gi-caption">条件 C の結果</div>
</div>

<div class="gi-item">

![w:400](assets/learning-curve.svg)

<div class="gi-caption">条件 D の結果</div>
</div>

</div>
```

### `figure` — 図の提示
<!-- 画像＋キャプション · 図を中心に見せるとき -->

```markdown
<!-- _class: figure -->

# 図解説スライド

![w:700](../assets/architecture.svg)

<div class="caption"><span class="fig-num">Fig. 1.</span> 提案手法のアーキテクチャ概要。入力系列を Embedding 後、N 層の Sparse Attention Block で処理し、タスクヘッドから最終出力を生成する。</div>

<div class="description">

- **入力層**: 生データを前処理し、特徴ベクトルに変換
- **中間層**: Sparse Attention Block を N 段積層
- **出力層**: タスク固有のヘッドで最終予測を生成

</div>
```

### `figure-cols` — 図＋考察の対比
<!-- 図 │ 解説 · 図と解説（観察・含意）を左右に並べるとき -->

```markdown
<!-- _class: cols-2 -->

# 図 + 解説（2カラム）

<div class="columns">
<div>

![w:440](https://via.placeholder.com/440x300/f0f2f5/1a1a2e?text=Result+Graph)

<div class="small muted center">Fig. 2: 学習曲線の比較</div>

</div>
<div>

## 観察

1. 提案手法（赤）は 50 epoch で収束
2. ベースライン（青）は 120 epoch 必要
3. 過学習の兆候は見られない

<div class="box-accent">

**収束速度 2.4x** — 事前学習の効果により初期段階から高精度

</div>

</div>
</div>
```

## メタ（meta）

### `title` — 始まり
<!-- 大タイトル · プレゼンの冒頭 -->

```markdown
<!-- _class: title -->

# スライドタイトル
## — サブタイトル —

著者名 $^{1}$, 共著者名 $^{2}$

$^{1}$ 所属大学 / 学部 &emsp; $^{2}$ 所属機関

Conference Name 2026 &ensp;|&ensp; 2026年 X月 X日
```

### `divider` — 転換
<!-- セクション区切り · 章の区切り -->

```markdown
<!-- _class: divider -->

# 1. セクション名

## このセクションの概要を一文で
```

### `summary` — まとめ
<!-- まとめリスト · 内容を要約するとき -->

```markdown
<!-- _class: summary -->

# まとめ

<ol class="summary-points">
<li>動的スパースマスクにより Attention 計算量を $O(n\sqrt{n})$ に削減</li>
<li>WikiText-103 で PPL 18.3 を達成し、標準 Transformer と同等精度</li>
<li>推論速度 2.4 倍、メモリ使用量 40% 削減</li>
<li>系列長 16K でも実用的な計算時間を維持</li>
</ol>
```

### `rq` — 問いの提示
<!-- 中央に問い · 研究質問を示すとき -->

```markdown
<!-- _class: rq -->

# Research Question

<div class="rq-main">
長系列データにおいて、精度を維持しつつ Attention の計算量を $O(n^2)$ 以下に削減できるか？
</div>

<div class="rq-sub">
— 動的スパースマスクによる適応的な注意範囲の選択
</div>
```

### `references` — 学術的裏付け
<!-- 文献リスト · 参考文献を示すとき -->

```markdown
<!-- _class: references -->

# References

<ol>
<li>
  <span class="author">Vaswani, A. et al.</span>
  <span class="title">"Attention Is All You Need."</span>
  <span class="venue">NeurIPS, 2017.</span>
</li>
<li>
  <span class="author">Devlin, J. et al.</span>
  <span class="title">"BERT: Pre-training of Deep Bidirectional Transformers for Language Understanding."</span>
  <span class="venue">NAACL-HLT, 2019.</span>
</li>
<li>
  <span class="author">Brown, T. B. et al.</span>
  <span class="title">"Language Models are Few-Shot Learners."</span>
  <span class="venue">NeurIPS, 2020.</span>
</li>
<li>
  <span class="author">He, K. et al.</span>
  <span class="title">"Deep Residual Learning for Image Recognition."</span>
  <span class="venue">CVPR, 2016.</span>
</li>
<li>
  <span class="author">Goodfellow, I. et al.</span>
  <span class="title">"Generative Adversarial Nets."</span>
  <span class="venue">NeurIPS, 2014.</span>
</li>
<li>
  <span class="author">Kingma, D. P. & Ba, J.</span>
  <span class="title">"Adam: A Method for Stochastic Optimization."</span>
  <span class="venue">ICLR, 2015.</span>
</li>
</ol>
```

### `appendix` — 補足資料
<!-- 補足 · 追加情報を示すとき -->

```markdown
<!-- _class: appendix -->

# ハイパーパラメータ一覧

<span class="appendix-label">Appendix A</span>

| パラメータ | 値 | 備考 |
|---|---|---|
| 隠れ次元 | 768 | BERT-base と同一 |
| ヘッド数 | 12 | |
| 層数 | 12 | |
| 学習率 | 3e-4 | cosine schedule |
| バッチサイズ | 64 | |
| Dropout | 0.1 | |
| スパース率 $k$ | $\sqrt{n}$ | 動的調整あり |
```

### `end` — 終わり
<!-- Thank You · プレゼンの終了 -->

```markdown
<!-- _class: end -->

# Thank you

Questions?

name@university.ac.jp
```

### `table-slide` — データ表示
<!-- 表 · 表形式でデータを示すとき -->

```markdown
<!-- _class: table-slide -->

# 表解説スライド

## 各手法の性能比較

| 手法 | Accuracy | F1 Score | 推論時間 (ms) | パラメータ数 |
|------|:--------:|:--------:|:------------:|:----------:|
| Baseline | 89.2% | 0.881 | 12 | 11M |
| Method A | 92.1% | 0.915 | 18 | 25M |
| Method B | 93.4% | 0.928 | 45 | 110M |
| **Ours** | **96.7%** | **0.962** | **15** | **18M** |

<div class="box-accent">

**提案手法の優位性**: 最高精度を達成しつつ、推論時間・パラメータ数ともに効率的。

</div>

<div class="footnote">すべての実験は同一のハードウェア環境 (NVIDIA A100) で実施</div>
```

