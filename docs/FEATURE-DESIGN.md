# marp-pptx 機能アーキテクチャ全体設計

> 2026-06-12 起草。tmu-cs / terminal-slide / git-presentation の機能評価に基づく。

## 1. エコシステムの中での位置づけ

| ツール | 入力 | 出力 | 強み |
|---|---|---|---|
| **marp-pptx**（本体） | Marp MD（52 型） | **編集可能 PPTX**（OMML 数式・実テーブル・ネイティブチャート） | 人間が PowerPoint で仕上げられる。決定論 lint・MCP/AI 連携 |
| [terminal-slide](https://github.com/lutelute/terminal-slide)（同作者） | MD / HTML | TUI・ブラウザ（アニメ制御・ギャラリー・エクスポート） | プレゼン実行系。動き・インタラクション |
| [git-presentation](https://github.com/lutelute/git-presentation)（同作者） | — | 自己完結 HTML | terminal-slide 製の作例。段階表示・動きの見本 |
| [marp-theme-tmu-cs](https://github.com/taishi-n/marp-theme-tmu-cs)（外部, MIT） | Marp MD + engine.mjs | HTML / PDF | **数式注釈・コードステップ強調・bib 引用など機能が marp-pptx より上位** |
| [marp-theme-dev](https://github.com/katsuzakitomohiro/marp-theme-dev)（外部, MIT） | Marp MD | HTML / PDF / 画像 PPTX | PowerPoint マスタとの設計契約。**→ テンプレ（`-p research` + 研究審査プリセット）として導入済み。これ以上の互換実装はしない** |

**marp-pptx の核は「編集可能 PPTX」**。HTML プレゼンの「動き」をそのまま持てない代わりに、
*クリック送り（スライド段階展開）* と *PowerPoint 上で編集できる実体*へ写像する。

## 2. 設計方針

1. **記法は tmu-cs 互換を正とする** — `[!step …]`・`[!math-annotate …]`・`[@key]`。
   同じ MD が tmu-cs（HTML/PDF）でも marp-pptx（編集可能 PPTX）でも意味を保って動くことを狙う。
2. **「動き」= ステップ展開** — tmu-cs の engine はステップごとにスライドを複製して段階表示する。
   marp-pptx も同じ戦略を採る：パース後に SlideData を複製し、PPTX ではクリック送りが「動き」になる。
   （PowerPoint ネイティブアニメーション XML 注入は将来オプション）
3. **劣化は明示的に** — HTML でしか成立しない機能（GIF 再生 UI・spectrogram・Kroki リモート図）は
   採用しないか、静的等価物（画像・脚注）へ落とす。落ちる場合は authoring warning で知らせる。

## 3. 機能インベントリと採用判断

| # | 機能（出典） | PPTX での表現 | 判断 |
|---|---|---|---|
| F1 | **数式注釈** `% [!math-annotate label= note= color=]`（tmu-cs） | 数式を行ごとに OMML/PNG で縦積み + 右に注釈カード + コネクタ線 | **Phase 1 で実装** |
| F2 | **コードステップ強調** `# [!step N action[:M]]`（tmu-cs） | 対象行に色帯+左バー、非対象行は減光。**step ごとにスライド自動複製（=動き）** | **Phase 1 で実装** |
| F3 | bib 引用 `[@key]` + `bibliography:`（tmu-cs） | 引用番号解決、スライド脚注、References 自動生成 | Phase 2 |
| F4 | セクションページ / TOC 自動生成（tmu-cs） | divider 自動採番・agenda 自動生成・スライド上部のセクション帯 | Phase 2 |
| F5 | 外部コード取り込み `[f.py](path)`（tmu-cs） | ファイル読込 → code 型へ展開 | Phase 2（小） |
| F6 | ギャラリービュー（terminal-slide） | 全スライドサムネの index — `render-gallery` が既にある（型カタログ）。デッキ単位のサムネ一覧出力に拡張 | Phase 3 |
| F7 | Kroki / mermaid 図（tmu-cs） | リモート依存。ローカル mermaid-cli があれば画像化 | Phase 3（任意） |
| F8 | GIF プレーヤー / spectrogram（tmu-cs） | PPTX では GIF 貼付で自動再生・音声は挿入のみ | 採用しない（静的等価で十分） |
| F9 | コード構文ハイライト（tmu-cs は Shiki） | pygments があれば run 単位で色分け | Phase 3（任意） |

## 4. アーキテクチャ

```
MD ──parse_slide──▶ SlideData（注釈・ステップを構造として保持）
        │                eq_annotations: [(tex, label, note, color)]
        │                code_steps:     [(line_idx, step, action, span)]
        ▼
   expand_step_slides()   ← パース直後の「デッキ変換」層（新設）
        │                step 番号の集合ごとに SlideData を複製し active_step を付与
        ▼
   PptxBuilder.build_*    ← active_step / eq_annotations を見て描き分け
        │                数式注釈: 行積み数式 + 注釈カード + コネクタ
        │                code: 行 paragraph 化 + 色帯 + 減光
        ▼
   編集可能 PPTX（OMML / 実テキスト / 実シェイプ）
```

- 指令の抽出は **parser**、見た目は **builder**、複製は **デッキ変換層** — 既存 52 型と直交し、
  どの型にも将来同じパターンで「指令 → 構造 → 描画」を足せる。
- 検証は従来どおり: pytest（リグレッション固定）+ devrender（soffice→PNG）+ visual_lint + collision。

## 5. フェーズ計画

- **Phase 1（今回）**: F1 数式注釈 + F2 コードステップ（展開含む）+ 機能ギャラリーデッキ
- **Phase 2**: F3 bib 引用、F4 セクション/TOC 自動化、F5 外部コード
- **Phase 3**: F6 デッキサムネギャラリー、F7 図、F9 構文ハイライト

## 6. 互換性ポリシー

- 指令なしの既存デッキは**一切変化しない**（指令検出時のみ新パスに入る）
- tmu-cs 記法のうち PPTX で意味を持たない属性（`:N` for math 等）は tmu-cs と同じく「parse して無視」
- すべての新機能にリグレッションテストを付け、`tests/test_regressions.py` の流儀で固定する
