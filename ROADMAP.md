# ROADMAP

marp-pptx の今後の拡張計画。

## 現状 (v0.1.0)

- CLI (`marp-pptx convert/types/preview/serve`)
- 45種のセマンティックなスライド型
- 10パレット + カスタムテーマ
- 数式: OMML (編集可能) + matplotlib PNG (fallback)
- Flask Web UI (アップロード → PPTX 変換)
- Docker イメージ
- pip パッケージ

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
