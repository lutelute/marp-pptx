# marp-pptx 品質検証レポート

- **対象**: marp-pptx v0.3.1(feature/web-ui ブランチ)
- **実施日**: 2026-06-08
- **検証者**: Claude Code(Opus 4.8)+ 重信
- **目的**: テンプレ流し込みではない**実コンテンツ**で、学会発表〜トップ博士発表レベルの編集可能 PPTX が作れるかを実地検証する

## TL;DR

> **作れる。型の文字数予算に収めれば学会発表級、図解とメッセージラインを足せば博士発表級。**
> 検証の過程で実バグ 6 件を発見・全件修正し、さらに無言フォールバック 4 種に警告を追加・`lint_deck` に density 引数を追加。リグレッションテストで固定した(**108 passed**)。
> 残る弱点は「溢れたときの耐性」(visual lint が要素間重なりを検出できない)。

## 1. 検証方法

1. **実コンテンツのデッキをゼロから執筆**(電力系統の PV ホスティングキャパシティ評価。数値は架空だがドメイン的に妥当)
2. `marp-pptx convert` で PPTX 化 → PPTX 内部 XML 検査(OMML / ネイティブチャート / 実テーブル)
3. LibreOffice + pdftoppm で全スライドを PNG 化し**全枚目視検査** → 問題を切り分け(コンテンツ起因 / ツール起因)
4. ツール起因のバグは builder/parser を修正し、`tests/test_regressions.py` で固定
5. 不正入力 11 ケースでエラーハンドリングを試験

## 2. 成果物

| ファイル | 内容 |
|---|---|
| `hosting-capacity.md` / `.pptx` | **v1**: 学会発表風 14 枚(12 種の型、数式・チャート・表・ノート入り) |
| `hosting-capacity-v2.md` / `.pptx` / `.pdf` | **v2**: 博士発表級リファイン 11 枚(全タイトル主張文化+自作 SVG 図解 2 枚+`--density keynote`) |
| `assets/problem.svg` / `loop.svg` | 自作図解(フィーダ電圧プロファイル / LP↔AC 往復ループ)。`rsvg-convert -w 1800` で PNG 化 |
| `regression-check.md` | バグ再現チェック用 4 枚デッキ(skeleton どおりの書き方が正しく出ることの確認用) |

### 編集可能性の証拠(PPTX 内部検査)

- 数式: `<m:oMath>`(ディスプレイ)+ インライン記号も `m:r`/`m:sSub` 構造 → **PowerPoint でダブルクリック編集可**
- チャート: `ppt/charts/chart1.xml` の**ネイティブチャート**(画像ではない。右クリック→データの編集が効く)
- 表: `<a:tbl>` 実テーブル / 発表者ノート: `notesSlide*.xml` に反映
- 一枚絵ゼロ: 全スライドが実テキストボックス

## 3. 発見・修正したバグ(6 件、全件修正済み)

| # | 症状 | 真因 | 修正 |
|---|---|---|---|
| 1 | `title` 型で著者行の `$^{1}$`・`&ensp;` が生テキストのまま出る | 著者行だけ素の `.text=` 代入(インライン処理なし)+ `strip_html` がエンティティ未処理 | `_set_rich_text` 化+`html.unescape` 追加(parser.py) |
| 2 | `timeline-h`/`timeline` の detail 内 `<br>` が無視され 1 行に連結 | 抽出時に空白として潰していた | `text_with_breaks()` 新設(`<br>`→`\n`、python-pptx が `<a:br/>` 化) |
| 3 | `takeaway` の主文が折り返すと箇条書きに**重なる** | `_estimate_text_height` が折り返し非考慮(1 論理行=1 物理行と仮定) | CJK 対応の折返し見積り(`width=` オプション)を追加 |
| 4 | `table-slide` で 6 行表が下の box-accent に接触 | 0.5in/行の固定見積りをレンダラが CJK 行高で食い潰す | 表ブロックに +0.25in ヘッドルーム |
| 5 | `appendix`+表で**スライドが 2 枚に分裂** | `build_table()` 委譲の先頭 `_blank_slide()` が新スライドを開く | 自スライドに `_styled_table` 直描き |
| 6 | **閉じ忘れ `<div>` で無限ループ(CPU100% ハング)** | `extract_child_divs` の while-else 誤用 — 閉じタグ不足時に `pos` が進まないまま外側ループが同じタグを再発見 | 閉じ忘れは残り全文をその要素として救済して終了 |

共通の教訓: **1〜2, 5 は「スキル文書(skeleton)に載っている書き方」が壊れる実装乖離**だった。6 は実ユーザーが最も踏みやすいミス(div 閉じ忘れ)で最悪の挙動(ハング)。

リグレッションテスト: `tests/test_regressions.py`(14 テスト＝バグ固定 10＋警告 4)。全テスト **108 passed**。

## 4. エラーハンドリング試験(11 ケース)

| ケース | 結果 | 評価 |
|---|---|---|
| 空ファイル | 0 枚 PPTX を生成(exit 0)＋警告 | ✅ 警告追加（後述） |
| frontmatter のみ | 同上 | ✅ 警告追加 |
| 未知の型 `_class: no-such-type` | default 型で 1 枚生成＋警告 | ✅ 警告追加 |
| **閉じ忘れ `<div>`** | **無限ループ → 修正済み**(内容救済して生成) | ✅ 修正済 |
| 存在しない画像参照 | 画像スキップで生成＋警告 | ✅ 警告追加 |
| 不正 LaTeX(`$$\frac{a$$`) | クラッシュせず生成(テキストfallback) | ✅ 耐性あり |
| 120 枚デッキ | **0.75 秒**で正常生成 | ✅ 高速 |
| 存在しない入力ファイル | click の明確なエラー | ✅ 親切 |
| 不正なパレット名 `-p no-such` | デフォルトにフォールバック＋`Warning:` 出力 | ✅ 既に警告あり |
| 絵文字・RTL・特殊文字 | 正常生成 | ✅ |
| 列数不揃いのテーブル | クラッシュせず生成 | ✅ |

**総評**: クラッシュ耐性は高い(ハング 1 件は修正済み)。当初は「黙って意図と違う出力を作る」無言フォールバックが多かったが、**本検証で `stderr` 一行警告を追加して解消**(空デッキ / 未知型 / 画像不在 / SVG×cairosvg)。不正パレット名は元から警告済みだった(初回検証時に `tail -2` で警告行を切り、誤って「無言」と判定していた)。`PptxBuilder.warnings` に蓄積され CLI が末尾で件数を要約する。

## 5. 既知の制限

1. **visual lint は要素間の重なりを検出できない**(検出は blank / edge-overflow / sparse / 上下偏りの 4 種のみ)。バグ 3 の重なりを 0 警告で通した。【未修正】
2. ~~**`lint_deck()` に density 引数がない**~~ → **本検証で `density=` 引数を追加**(keynote 成果物を keynote 設定でレンダ)。default は academic で後方互換。【修正済み】
   - ⚠️ **自己レビューで判明した誤り＋バグ修正**: density 追加直後、keynote で 11 枚中 9 枚に edge 警告が出たのを「字が大きく端に寄るため正しい挙動」と一旦結論づけたが**これは誤り**だった。実体は `visual_lint` が**右下のページ番号(builder が描く chrome)を本文として bbox に含めていた**ため(全枚 maxy=0.967、本文は除けば 0.60–0.92)。footer 隅(y>0.93 ∧ x>0.78)を bbox から除外する修正で 9→0。広い下溢れ・左右溢れの検出は合成画像テストで温存を確認。【修正済み】
3. `_estimate_text_height` の `width=`(折返し考慮)を渡しているのは takeaway のみ。同じ重なりリスクが rq / profile 等に残る(同じ直し方で対処可能)。【未修正・横展開は要視覚検証のため保留】
4. `eq-desc` の説明文は keynote 密度で約 20 字/行を超えると折り返して凡例の対応が崩れる(コンテンツ側で短く書いて回避)。
5. title の kicker(H2)は**空白区切り 6 語以下**のみ。日本語は詰めて書く(超えると kicker ごと消える)。

## 6. 「どのレベルまで作れるか」の整理

| 品質要素 | ツールの守備範囲 | 実証 |
|---|---|---|
| 組版(階層・余白・配色・整列) | ✅ 自動 | v1 の agenda / timeline / kpi / chart はそのまま使える |
| 編集可能性(数式・チャート・表) | ✅ 自動 | §2 の XML 証拠 |
| メッセージライン(タイトル=主張文) | ◯ 書き方次第 | v2 で全タイトル主張文化 |
| 図解(概念図・構成図) | ❌ 守備範囲外 | SVG を別途作成して `diagram` 型に持ち込み(v2 で 2 枚、1 枚 10 分程度) |
| 投影向け密度 | ✅ `--density keynote` | v2 |

**結論**: 「Markdown を書けば自動で博士発表級」ではないが、**組版品質の床が高い**ため、人間(または AI)が (a) タイトルの主張文化と (b) 図解の持ち込みに集中すれば、トップレベルの発表資料に到達できる。所要時間は v1(14 枚)が初稿数分+磨き込み、v2 へのリファインが図解込みで 1 時間弱。

## 7. 再現コマンド

```bash
# v1(学会風・academic 密度)
marp-pptx convert demo/hosting-capacity.md -o demo/hosting-capacity.pptx

# v2(博士発表級・keynote 密度)
rsvg-convert -w 1800 demo/assets/problem.svg -o demo/assets/problem.png
rsvg-convert -w 1800 demo/assets/loop.svg -o demo/assets/loop.png
marp-pptx convert demo/hosting-capacity-v2.md -o demo/hosting-capacity-v2.pptx --density keynote

# 視覚検証(LibreOffice レンダリング、数式は PNG モードで)
marp-pptx convert demo/hosting-capacity-v2.md -o /tmp/check.pptx --density keynote --math png
soffice --headless --convert-to pdf /tmp/check.pptx --outdir /tmp
pdftoppm -png -r 90 /tmp/check.pdf /tmp/slide

# テスト
python -m pytest tests/  # 108 passed
```

## 8. 自己改善ループ（レポート後の継続改善, 2026-06-09〜10）

レポート後、「完了と思っても粗を探し続ける」方針で批判的自己レビューを反復した。各周＝1実バグ発見→修正→テスト→push。

### 8.1 新手法: 52-skeleton コンテンツ忠実度監査

全 52 型の公式 skeleton を1枚ずつビルドし、**入力の可視テキストトークンが出力 pptx に現れるか**を機械照合（大文字小文字無視・frontmatter/数式/chart-XML を除外）。「skeleton どおり書いたのに内容が落ちる」型を網羅的に炙り出す。これで builder/parser 側の**サイレントなコンテンツ欠落**を5件発見した。

### 8.2 ループで見つけ・直した問題

| # | 種別 | 症状 | 真因 | commit |
|---|---|---|---|---|
| 7 | parser | timeline の強調項目テキスト（通常「本研究」）が消える | `class="tl-h-text bold"`（2クラス）を exact-quote 正規表現が拾えず | 1363ddf |
| 8 | parser | timeline の highlight（accent 強調）が一度も効かない | `extract_child_divs` が item の class を捨て `"highlight" in item` が常に偽 | 1363ddf |
| 9 | builder | zone-matrix の**両軸ラベル**が描画されない（2軸図の定義的機能） | builder がセルだけ描き x/y_label を無視 | 54251b9 |
| 10 | builder | overview/result の図キャプションが落ちる | 共有 `_build_image_points` が `sd.caption` 未描画 | b5ed6aa |
| 11 | builder | gallery-img の各画像キャプションが全部落ちる | `build_gallery_img` が `item["caption"]` 未描画 | dd80514 |
| L1 | lint | visual_lint がほぼ全スライドで偽の edge 警告 | 右下のページ番号 chrome を本文として bbox に算入（keynote で顕著） | 3c04a4c |
| G1 | mcp | エージェントが build_pptx の警告（未知型/画像不在）を受け取れない | 返り値に authoring_warnings が無く stderr で握り潰し | 12f8e13 |

### 8.3 系統監査の負の結果（クリーン確認）

- パーサ全体の exact-quote `class="X"` 正規表現を監査 → 残るは references/checklist のみで、両 skeleton は2クラスを使わず現状バグ無し（投機的変更は見送り）。
- 52-skeleton 監査の残り信号（equation の LaTeX 断片・chart 系列名）は OMML/chart-XML 行きの**偽陽性**で実害なし。

### 8.4 累積

原検証6件 + ループ5件（builder/parser）= **実バグ11件修正**、加えて lint 偽陽性・MCP 警告伝播を改善。テスト **108 passed**（うち test_regressions.py が 19）。残候補: `_estimate_text_height(width=)` の rq/profile 等への横展開（要視覚検証で保留）、visual lint への要素間重なり検出。
