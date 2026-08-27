# ROADMAP

marp-pptx の今後の拡張計画。

## v0.4（完了）: 実測レイアウト精度

「資料としての精度」を推測から実測に置き換えた。

- **`marp_pptx.metrics`** — 実フォントの字送り（FreeType）で幅・折り返し（禁則処理つき）・
  高さ・収まるサイズを計算。ビルダー / `visuallint` の衝突検出 / `doctor` が**同じ物差し**を使う
- **`marp_pptx.audit`** — 組み上がった PPTX の幾何を検査（溢れ・重なり・スライド外・
  WCAG AA・書体・レイアウト単調）。レンダラ不要、数十 ms
- **`marp_pptx.pkgcheck`** — OOXML パッケージ整合（参照切れ・content-type・拡張子詐称）
- **`marp-pptx doctor`** / MCP **`check_deck`** / `build_pptx` の `defects` /
  `from-paper` の自動修復ループ
- ビルダー側の実バグ修正: 番号付きリストのインライン記法未処理、カラム高さの推定破棄、
  行間を無視した高さ予約（title/definition/big-number ほか）、カード chrome の未計上 など
- **同梱 47 テンプレート + プリセット × 5 テーマで error 0** をテストで固定

残る既知の警告: `claude` / `research` パレットの一部で小さい白抜き文字が AA 未満（3.0:1）。
配色側の調整が必要（`accent_text` の導入で本文側は解決済み）。

## ★ 最重要目標: PPTX ⇄ MD 双方向ラウンドトリップによる品質向上

**ビジョン**:
世の中の**良質な学会発表 PPTX** を学習データとし、
`PPTX → 構造抽出 → Marp MD → PPTX 再現` の双方向変換を通じて
見た目の品質とテンプレ語彙を継続的に改善する。

```
 良いPPTX(実サンプル)
      │
      ▼ ①構造抽出 (pptx2md + レイヤ分類)
 中間表現 (スライド型 / テキスト / 数式 / 画像 / 配置)
      │
      ▼ ②Marp化 (既存49型にマッピング, 不足型は新設)
 Marp MD (学習データ / テンプレ源)
      │
      ▼ ③Builder (PptxBuilder)
 再現 PPTX
      │
      ▼ 比較評価 (見た目差分 / 文章数式の対応維持)
 フィードバック → ① or ②の改善
```

### 非交渉原則
1. **見た目が重要**: 学会発表として違和感ない配置・余白・配色を最優先
2. **文章と数式がバラバラになってはいけない**:
   - 数式とその直前/直後の文章は必ず同じスライド・同じ論理ブロック内に保つ
   - 「変数説明」「結果の式」と本文が分断される出力は **不合格**
3. **レイヤ分解は許容**: 背景/タイトル/本文/数式/図をレイヤ単位で扱ってよい (編集性維持の前提)
4. **編集可能性を保つ**: 再現PPTXも OMML 数式 + 実テキストボックスを維持

### 進捗ゲート (この順で改善する)
| 段階 | 達成条件 |
|------|---------|
| R1 | 単一スライドの pptx→md→pptx でテキスト100%復元 (`pptx2md` 既存) |
| R2 | レイアウトクラス (`title/cols-2/equation/...`) の自動推定が実サンプル70%で一致 |
| R3 | 数式と本文の「論理ペア」(例: 「式+変数説明」) が分離しない |
| R4 | 学習データ全49型に対し再現PPTXが原本と視覚差 <10% (IoU/SSIM ベース) |
| R5 | 未知レイアウト検出 → 新スライド型を半自動提案 |

### 学習データ
- `learning-data/slide-templates/` (.gitignore 済) に良質テンプレを収集
- スコアリング rubric は `learning-data/rubric.md` を参照 (70点未満は除外)

---

## 現状 (v0.1.0)

- CLI (`marp-pptx convert/types/preview/serve`)
- 45種のセマンティックなスライド型
- 10パレット + カスタムテーマ
- 数式: OMML (編集可能) + matplotlib PNG (fallback)
- Flask Web UI (アップロード → PPTX 変換)
- Docker イメージ
- pip パッケージ

## v0.1.5: デザイントークン化 + テーマプラグインシステム（C 案）

**背景**: 「Claude design」のような統一感ある美しいスライドを実現したいが、特定ブランドに依存させると他ユーザーが使えない OSS にならない。そこで「美しさを実現する**仕組み**」を汎用化し、テーマは差し替え可能にする。

**方針 (C 案)**:
- デザイントークン（color / type / spacing / elevation）を `ThemeConfig` に集約
- テーマは JSON ファイル化、`themes/` 配下にプラグイン形式で配置
- 公開リポには中立サンプル (`minimal`, `academic`, `presentation`) を同梱
- 「Claude inspired」等のブランド色テーマはユーザー側で private 配置（gitignore 推奨）

**現状把握** (調査済):
- `builder.py` 2,150 行、`Pt()` 141 箇所 + `Inches()` 155 箇所 = 約 300 リテラル散在
- 46 個の `build_*` メソッドが独立にリテラルを埋め込み（共通化されていない）
- `_fs()` ヘルパーは存在（font_scale 対応）、`margin_scale` は定義済みだが未使用

**Phase 分割**:
| Phase | 内容 | 工数 |
|---|---|---|
| P1 | `_ms()` / `_inches()` ヘルパー追加 + 1-2 型でパイロット | 1-2 日 |
| P2 | テーマ JSON スキーマ確定 + 中立サンプル 2-3 種 | 2-3 日 |
| P3 | 46 型を機械的に token 化（型ごとにテスト） | 5-7 日 |
| P4 | 中間 JSON 経路 + HTML プレビュー（編集可能性維持） | 1 週間 |

**非交渉原則** (本 ROADMAP 冒頭原則と整合):
- pptx には SVG を一枚絵として埋め込まない（編集可能性を必ず維持）
- HTML/SVG プレビューは「見た目確認用」のみ、出力は OMML + テキストボックス
- 数式と本文の論理ペアは絶対に分離させない（既存原則）

**v0.2 との関係**: v0.2 の `font_scale` / `margin_scale` スライダーは v0.1.5 P1-P3 完了が前提。

---

## v0.2: プレビュー調整 UI（① "Preview Adjustment"）

**目的**: MD → 調整 → PPTX の往復を1画面で完結させる。

**実装予定**:
- Web UIに `/preview` 画面を追加
- アップロードした MD をパースし、スライドごとに：
  - 型 (slide_class)
  - H1 / H2
  - 本文の文字数・bullet数
  - 含まれる画像パス
  - 警告 (例: 未知の型、画像不在)
- グローバル設定スライダー:
  - パレット選択
  - **font_scale** (0.7 - 1.3): 全フォントを一括スケール
  - **margin_scale** (0.7 - 1.3): マージンを一括スケール
  - spacing (compact/normal/generous)
- 設定を JSON 形式で保存/読込
- 「PPTX 生成」ボタンで最終出力

**前提条件**:
- `ThemeConfig` に `font_scale`, `margin_scale` フィールド追加
- `PptxBuilder` のフォント/マージン適用箇所に scale 乗算を組み込む
  - 現状 `Pt(N)` リテラルが builder 全体に散在 → `self._pt(N)` のような helper に置換が必要

---

## v0.3: WYSIWYG エディタ（② "Visual Editor"）

**目的**: ブラウザ上でスライドを直接ドラッグ/リサイズして PPTX に反映。

**設計**:
```
MD → 中間JSON (型・位置・サイズ・色をパラメータ化)
     ↓
ブラウザエディタ (React/Svelte + Canvas or DOM drag-drop)
     ↓
修正後JSON → PptxBuilder → PPTX
```

**中間JSON例**:
```json
{
  "slides": [
    {
      "type": "funnel",
      "h1": "採用プロセス",
      "items": [
        {"label": "応募", "value": "1,000", "size_override": 1.2},
        ...
      ],
      "layout_override": {"margin_l": 0.5, "font_scale": 0.9}
    }
  ],
  "global": {
    "palette": "navy",
    "font_scale": 1.0
  }
}
```

**工数見積**: 1-2週間
- フロントエンド新規開発 (React/Svelte + Vite)
- JSON スキーマ設計
- Builder の JSON 直読み経路追加
- 双方向変換 (MD ↔ JSON)

**いつやるべきか**:
- 複数人で同じテンプレートを使い回す必要が出た時
- ピクセル単位の微調整が頻繁に必要になった時
- v0.2 の範囲内（スライダー調整）で不足を感じ始めた時

**注意点**:
- ツールの本質は「型を選べば正しいレイアウト」。
- ②を入れると「型 + 個別調整」が常態化し、型システムが形骸化するリスクがある。
- 自由度は「型を増やす」方向で吸収できないか、まず検討する。

---

## その他の小さな拡張

### v0.2 に含めるかもしれない
- 数式の matplotlib → LaTeX フォールバック (`usetex=True`)
- SVG 入力サポートの強化 (cairosvg 依存を必須化)
- `marp-pptx init` コマンド: サンプル MD を生成

### 将来の検討
- PowerPoint アニメーション対応 (現状は静的スライド)
- スピーカーノート対応 (Marp の `<!-- note -->` → PPTX notes slide)
- リアルタイム共同編集 (WebSocket + 中間JSON)
- AI による型推論 (無指定時に内容から `<!-- _class: -->` を自動補完)

---

## 意思決定メモ

### なぜ ① を先にやるか

- 実装コスト低 (2-3日 vs 1-2週間)
- 既存の Flask UI と型システムに自然に乗る
- 型の枠内での調整に留まるので、ツールの哲学を壊さない
- ①のJSON状態保存は ② への下位互換になる

### なぜ完全 Python 化したか

- Node.js 依存を消して 500MB 削減
- pip 一発インストール可能に
- クロスプラットフォーム (Windows/Linux/macOS) の確実性
- LLM 統合は後付け可能な設計にしてある

---

## 動きとインタラクティブ化（2026-08-27 計画）

「GIF もインタラクティブ化も入ってくる予定」に向けた PPTX の実力マップと段階計画。

### 現状すでに動くもの（実証済み）

- **アニメーション GIF**: `![w:400](anim.gif)` で `ppt/media/*.gif` にそのまま埋め込み、
  スライドショーで自動再生される（`_image_or_placeholder` 経由・欠像時プレースホルダも有効）
- **クリック送りの「動き」**: `# [!step N action:M]` が code 型をステップごとに
  スライド複製（tmu-cs 互換）。実質クリックアニメ
- **内部ナビゲーション**: agenda の各項目が対応する divider へ
  `ppaction://hlinksldjump` でジャンプ（`_link_agenda_sections`、i 番目の項目 → i 番目の divider）

### 次の段階（実装順）

1. **`<p:timing>` ネイティブアニメ注入** — step 展開を「1枚でクリック出現」に置換する
   オプション（`--steps native`）。appear/fade をビルド時に XML 注入。スライド複製版は
   フォールバックとして維持
2. **ナビゲーション拡張** — divider から agenda へ戻るリンク、footer_bar のセル click、
   appendix への「詳細は付録」ジャンプ
3. **動画埋め込み** — `![w:800](demo.mp4)` → `p:video`（poster 画像自動生成）。
   GIF と同じ経路で `_resolve_image` を拡張
4. **セクションズーム / サマリーズーム** — PowerPoint の Zoom 機能（要 XML 調査。
   壊れやすいので validate.py 系の検査を先に用意）

### やらないこと（PPTX の外）

- VBA / マクロ（.pptm の領分）・Python 実行デモ → terminal-slide（Pyodide）担当
- ブラウザ実行前提の GIF 再生 UI・spectrogram（HTML 専用と判定済み、FEATURE-DESIGN.md）
