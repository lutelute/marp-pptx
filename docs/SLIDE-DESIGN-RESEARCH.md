All token values from the research are confirmed against the actual code (MARGIN_L=0.62in, SZ_TITLE=28/H2=22/BODY=18pt, LINE_BODY=1.30, claude palette cream #faf9f5 + clay #d97757 + ink #141413). I have everything needed to write the grounded report.

# marp-pptx 設計改善のための統合リサーチレポート

8観点(学会発表/Presentation Zen・TED/コンサル/キーノート・プロダクト/タイポグラフィ・グリッド/データ可視化/色覚・アクセシビリティ/ギャラリー)を統合した最終版。重複は単一原則に集約し、定量値と実在確認済み出典URLのみを併記する。コード照合済み:`layout.py` の `MARGIN_L=0.62in / SZ_TITLE=28 / SZ_H2=22 / SZ_BODY=18 / SZ_COL=16 / SZ_SMALL=14 / SZ_ZONE_B=15pt / LINE_BODY=1.30`、`data/themes/palettes/claude.css` の cream `#faf9f5` + clay `#d97757` + ink `#141413` を確認。

---

## 1. 優れたスライドの核心原則(横断)

8観点すべてで反復・相互裏付けされた原則を、定量値と出典つきで提示する。一次/精読確認できたものを優先採用した。

### 原則1 — 1スライド=1メッセージ
全8観点で完全一致した最上位原則。複数の論点を持つなら分割する。タイトルに「and / および / 、」が現れたら分割サイン。
- TED公式「No slide should support more than one point」: https://www.ted.com/pages/create-prepare-slides
- PLOS Comp Biol Rule 1「one central objective」: https://journals.plos.org/ploscompbiol/article?id=10.1371/journal.pcbi.1009554
- コンサル「点が2つなら2枚」+ 60秒ルール: https://deckary.com/blog/consulting-slide-standards

### 原則2 — 1スライドの要素数 ≤ 6、本文は ≤ 6行/各行 ≤ 約40字(英語6語)
査読付きの定量規範。要素=テキストブロック・図・ラベル・キャプション等の独立視覚単位の合計。
- PLOS Rule 7「total number of elements ≤ 6」: 上記PLOS URL
- Peyton Jones「Six or seven 'things' is quite enough」: https://www.microsoft.com/en-us/research/wp-content/uploads/2016/08/giving-a-talk.pdf
- 6×6ルール(≤6項目・各≤6語): https://www.free-power-point-templates.com/articles/the-6x6-rule-in-presentation-design/
- Duarte の上限「≤30語・≤5論点、75語超は文書(document)であって投影物ではない」: https://www.duarte.com/resources/books/slideology/(本体ページ未精読・数値は複数二次要約で一致)

### 原則3 — 1スライド ≒ 1分(20分発表 ≒ 20枚)
情報密度の上限を時間軸に変換する規律。
- PLOS Rule 2「a 20-minute presentation should have around 20 slides」: 上記PLOS URL

### 原則4 — 見出しは「主張文(assertion / action title)」にする
名詞ラベル(「結果」「Results」「Method」)でなく、結論を完全文で。可能なら数値を含める(例「Reported cases decreased by 17%」「提案手法が誤差を17%削減」)。Assertion-Evidence群は対照実験で理解・記憶が有意に高かった(p<.01)。コンサル流の定量規律は **≤15語・2行以内**(典型8語未満)、能動態。見出しを上から読むだけで物語が通る(横の論理 / Horizontal Flow)。
- PLOS Rule 3 & 8(自己完結スライド): 上記PLOS URL
- Alley & Neeley(Assertion-Evidence): https://writing.engr.psu.edu/research.html
- アクションタイトル(15語/2行): https://slideworks.io/resources/how-to-write-action-titles-like-mckinsey
- 横の論理: https://slideworks.io/resources/how-mckinsey-consultants-make-presentations
- データ図の主張型タイトル: https://sixminutes.dlugan.com/slide-charts/

### 原則5 — フォントは大きく、サイズ種類 ≤ 3・書体 ≤ 2、下限は会場で決まる
口頭発表の本文下限は用途依存:配布/PC閲覧 20pt+、会議室 24pt+、大会場 30pt+。学会本文は最低 18–24pt。サイズは1スライド3種以内、書体は1–2種、サンセリフ基調、強調は太字(イタリック・下線・全大文字は避ける)。
- 用途別下限(配布20/会議室24/大会場30pt): https://thepopp.com/how-big-should-text-be-in-powerpoint/
- キーノート本文28pt未満不可・48pt+目標、サイズ≤3種: https://www.brightcarbon.com/blog/presentation-font-size/
- 学会アクセシビリティ(本文24pt+/見出し32pt+): 上記 brightcarbon / PLOS Rule 7
- ジャンプ率(タイトル÷本文)1.2–2.0×(典型1.3×): https://info.winschool.jp/detail46/
  - 現状照合:`28/18≈1.56`、`22/18≈1.22` はジャンプ率の推奨帯に合致。一方で **body 18pt は配布の下限(20pt)をわずかに下回り、大会場(30pt)には明確に不足**。

### 原則6 — 型スケールは単一比率(モジュラースケール)で統一
ベース×固定比率で各階層を導くとリズムと階層が安定。実用比率:Minor Third 1.2 / Major Third 1.25 / Perfect Fourth 1.333 / Perfect Fifth 1.5。
- モジュラースケール早見: https://spec.fm/specifics/type-scale
  - 現状照合:body18→H2 22→title28 は比率が `1.22 / 1.27` と**非統一**。base18・r=1.333 で統一すると body18→H2 24→title32(さらに caption 14)と整然。

### 原則7 — 行長 45–75字(上限80字)、行間 1.2–1.5倍
古典タイポグラフィ(Bringhurst/Ruder)とUX研究が一致。本文1行は45–75字(理想66字付近)、80字超で疲労、100字で破綻。行間はサイズの1.2–1.5倍(本文長文は1.4–1.5、見出しは1.1–1.2)。CJKは全角約33–40字/行が目安。
- 行長: https://practicaltypography.com/line-length.html / https://baymard.com/blog/line-length-readability
- WCAG 1.4.8(行長 ≤80字): https://www.w3.org/WAI/WCAG21/Understanding/contrast-minimum.html(同規格群)
  - 現状照合:`LINE_BODY=1.30` は推奨帯(1.2–1.5)の下寄りで妥当。ただし長文型(definition/quote/summary/overview)では 1.4–1.5 へ寄せると可読性向上。

### 原則8 — コントラスト比:本文 4.5:1 以上、大文字(≥18pt or ≥14pt太字)3:1 以上
WCAG 2.1 AA。投影はコントラストを失うため床値であって目標ではない。
- W3C WCAG 1.4.3: https://www.w3.org/WAI/WCAG21/Understanding/contrast-minimum.html
- WebAIM Contrast Checker(検証用): https://webaim.org/resources/contrastchecker/
  - 現状照合:**ink `#141413` on cream `#faf9f5` ≈ 18:1**(AAA余裕クリア)。**clay `#d97757` on cream ≈ 3.0:1** で本文には不可、大見出し/アクセント線/大数字(3:1基準)専用に限定すべき。

### 原則9 — 余白(ホワイトスペース)を機能として使い、Signal/Noise比を最大化
装飾・不要ロゴ・フッター・グリッド線・3D効果・チャートジャンクを削る。「より少ないことで足りるなら、多くを使うのは無駄」(Tufte)。具体的余白%の権威ある一次数値は未確認(定性原則)。
- Garr Reynolds: https://www.garrreynolds.com/design-tips
- IEEE ProComm「SNRは原則であってルールではない(盲従するな)」: https://procomm.ieee.org/design-your-slides-to-maximize-signal-to-noise-ratio/
  - 現状照合:`margin 0.62in`(SWの約4.6%)は「最小余白≒タイトル文字高(title28→約0.39in)」を上回りセーフ域内。キーノート用途では 0.9in 級への拡大余地あり。

### 原則10 — 画像優位性・図を必須にする(テキストのみスライドを避ける)
画像は単語より記憶に残る(Shepard 1967:認識 画像98% vs 単語88%)。論文図はそのまま貼らず1スライド1パネルに再描画、フォント拡大・線を太く、出典は「modified from [source]」。テキストのみのスライドはほぼ作らない。
- 画像優位性: https://en.wikipedia.org/wiki/Picture_superiority_effect
- PLOS Rule 6「almost never slides that only contain text」: 上記PLOS URL
- 図の再設計(before/after): https://mitcommlab.mit.edu/cee/2024/01/15/redesigning-existing-figures-for-slides/

### 原則11 — 1アクセント色規律(60-30-10)、色数 ≤ 3–5
60%地色/30%文字/10%アクセント。前注意属性(色・サイズ・位置)は「狙って・控えめに」。
- Garr Reynolds(色2-3): 上記URL / Duarte コントラスト=単一焦点: https://www.duarte.com/blog/ultimate-guide-to-contrast/
  - 現状照合:ツールの「1アクセント色規律」は本原則と完全整合。

### 原則12(データ図) — グレー地+1色強調、円より棒、ミニマル表
強調は全データの10–20%以内・単純図は1点・中複雑図は5点まで。比較は長さ(棒)で:棒は円より約1.96倍正確(Cleveland 1984)、応答も速い。ただし「割合推定」タスクなら円も許容(タスク依存)。表は縦罫廃止・横細罫のみ、数値は右揃え/小数点揃え・等幅。
- 1点強調・muted/highlight: https://sixminutes.dlugan.com/slide-charts/
- 円vs棒(タスク別): https://journals.sagepub.com/doi/10.1177/14738716241259432
- 表設計(縦罫廃止・細罫・数値揃え・行長50–70字): https://www.csescienceeditor.org/article/best-practices-in-table-design/

### 原則13(データ図) — 色覚対応:赤緑を隣接させない、Okabe-Ito 8色、冗長符号化
赤緑色覚異常は男性最大8%・女性0.5%。安全は青`#0072B2`×橙`#E69F00`、または緑×マゼンタ。カテゴリ配色は Okabe-Ito 8色(`#E69F00 #56B4E9 #009E73 #F0E442 #0072B2 #D55E00 #CC79A7 #000000`)を既定に(グレースケール印刷でも判別可)。色だけに依存せず形・線種・パターン・直接ラベルを併用。
- Bang Wong, Nature Methods 2011: https://www.nature.com/articles/nmeth.1618
- JFLY / Color Universal Design: https://jfly.uni-koeln.de/color/
- Okabe-Ito hex一覧: https://conceptviz.app/blog/okabe-ito-palette-hex-codes-complete-reference

### 原則14 — 目次で始めない/例から入る、戦略的「間」を使う(純白回避)
研究トークは目次スライドで始めず動機づけの「例」から入る(中盤約1/3地点に整理スライド1枚は有効)。戦略的な空白スライドで注意を話者へ戻すが、純白は機材故障と誤認されるため地色を維持。
- Peyton Jones §2「jump straight in with an example」「目次禁止」: 上記Peyton Jones PDF
- 空白スライド戦略: https://ideas.ted.com/6-dos-and-donts-for-next-level-slides-from-a-ted-presentation-expert/
- 高橋メソッド(巨大文字・1メッセージ): https://ja.wikipedia.org/wiki/高橋メソッド

---

## 2. 学会発表 と ビジネス/キーノート の差分

### 共通(両用途で守るべき不変原則)
- 原則1(1スライド1メッセージ)、原則4(主張文/アクションタイトル)、原則5(サイズ≤3種・書体≤2)、原則8(WCAGコントラスト)、原則9(余白・SNR)、原則11(1アクセント色)、原則14(純白回避・1スライド1焦点)はすべて両用途で一致。特に「見出しに結論を書く」は学会(Assertion-Evidence)とコンサル(Action Title)が独立に同じ結論へ到達している。

### 違い(用途で最適値が分かれる軸)

| 軸 | 学会・研究発表 | ビジネス/キーノート | 出典 |
|---|---|---|---|
| 本文フォント下限 | 18–24pt(配布兼用・情報密度高) | 28pt+(理想48pt+、投影前提) | brightcarbon / thepopp |
| 1スライドのテキスト量 | 図1点+補足1行×数本まで許容 | 1–2行(Apple実績で全体≈19語/10–12枚) | PLOS Rule 6 / crappypresentations |
| 余白 | やや高密度(現行 0.62in 維持可) | 広め(0.9in級・要素≤4) | Garr Reynolds |
| 行間 | 1.30前後(箇条書き多) | 1.4–1.5(長文・呼吸重視) | Baymard |
| 表/データ | 5–7行・変数1–3個(査読由来) | 単一の大数字を1枚で見せる(big-number) | PMC / piktochart |
| 構成の起点 | 動機づけの例 → RQ → 手法 → 結果 → takeaway(目次で始めない) | SCQA → ピラミッド頂点(結論+3論拠)のExecutive Summary | Peyton Jones / think-cell |
| 出典脚注 | 全データ系で必須(末尾に参照スライド) | 全データ系で必須(下端small-caps) | Sheridan / Slideworks |
| 背景 | 明背景+濃文字(配布・明室前提) | 暗転(ink背景×cream文字)を章転換・大数字で演出可 | PLOS Rule / SlideGenius |

要するに「**学会=高密度・配布兼用・証拠重視/キーノート=低密度・投影専用・1焦点演出**」。同じ49型を**密度プリセット(academic / keynote)で切り替える**のが最も合理的な設計解。

---

## 3. marp-pptx への具体的改善提案(実装粒度)

実装着手点(絶対パス):
- `/Users/shigenoburyuto/Documents/GitHub/marp-pptx/src/marp_pptx/layout.py`(トークン)
- `/Users/shigenoburyuto/Documents/GitHub/marp-pptx/src/marp_pptx/builder.py`(型実装・テキスト流し込み・lint追加先)
- `/Users/shigenoburyuto/Documents/GitHub/marp-pptx/src/marp_pptx/theme.py`(パレット・コントラスト検証)
- `/Users/shigenoburyuto/Documents/GitHub/marp-pptx/src/marp_pptx/cli.py`(プロファイル/フラグ)
- `/Users/shigenoburyuto/Documents/GitHub/marp-pptx/src/marp_pptx/data/themes/palettes/`(配色・clay制約)

### High(効果大・実装容易・全観点で要請)

**H1. 本文量バジェットのlinter警告(原則1/2/3)**
builder のテキスト流し込み後に各スライドの「要素数・本文行数・1行字数」を集計し、閾値超で stderr 警告(エラーではなくwarning)。
- 要素数 > 6 → `[overflow] slide N: 9 elements (>6)`
- 本文 > 6行 または 1行 > 40字(CJK全角=2カウント、全角約40字)→ 警告
- 語数 > 75(英語)→ strong-warning「文書化している。分割を検討」
- `cols-2/3 / zone-* / card-grid` は合計要素数で集計、各カードは ≤2行推奨
- 出典:PLOS Rule 7 / 6×6 / Duarte。実装はカウント関数1本+型別閾値テーブル。

**H2. clay `#d97757` をテキスト用途から外す制約をコード化(原則8/13)**
cream上で約3.0:1のため、本文・キャプションへの clay 適用を builder で禁止。アクセント線・大見出し(3:1で可)・図形塗り・大数字には許可。必要なら本文用にダーク化した clay 派生色を自動生成。出典:WCAG 1.4.3。`theme.py` + 各型レンダラ。

**H3. 全テーマ/全パレットに contrast-lint を組込む(原則8)**
`theme.py` で文字色×地色・accent×地色を WCAG式で計算し、本文 < 4.5:1・大見出し/ラベル < 3:1 でビルド時 warning + 候補色提案。`--strict-contrast` でエラー昇格。claude/minimal/academic10 全パレットを検査。出典:WCAG / WebAIM。

**H4. 用途プリセット `--density academic|keynote`(またはfront-matter)(原則5/7/9・第2章)**
1フラグで以下をまとめて切替:
- academic:body18 / line 1.30 / margin 0.62in / 要素上限6 / 本文行上限6(現状維持=既定)
- keynote:body24–28 / title36–40 / line 1.45 / margin ~0.9in / 要素上限4 / 本文行上限2–3
H1の警告閾値もこのプリセットに連動。Duarte Slidedocs の「投影用 vs 読む用」二態を型レベルで実装。出典:brightcarbon / thepopp / Garr Reynolds。`cli.py` + `layout.py` をプロファイル化。

**H5. 見出しの主張型(assertion / action-title)lint(原則4)**
H2/H1が名詞句・体言止め・汎用ラベル(「結果」「考察」「Results」「Method」)の場合に info警告「結論を1文で(例:提案手法が誤差を17%削減)」。`result/multi-result/kpi/figure/table/chart` 型では特に「数値inタイトル」を推奨。コンサル規律の15語・2行超も警告。出典:PLOS Rule 3 / Slideworks / six minutes。

### Med(設計改善・新型・効果中)

**M1. 型スケールを単一比率に再設計(原則6)**
body18・r=1.333(Perfect Fourth)で統一:body18 → H2 24 → title32、caption14/16 を追加。`r` をテーマトークン化し academic系でも一貫。現状の `22/1.27` 不揃いを解消。出典:spec.fm。`layout.py`。

**M2. line height を文脈別トークンに分岐(原則7)**
一律 `LINE_BODY=1.30` を分割:`line_tight=1.10`(title/H2/ラベル)、`line_list=1.30`(箇条書き)、`line_prose=1.45`(definition/quote/summary/overview の長文)。出典:Baymard。`layout.py` + 各型レンダラ。

**M3. `chart` 型を新設 + Okabe-Ito 安全パレット内蔵(原則12/13)**
折れ線/棒用 `chart` 型を追加。系列色デフォルトを Okabe-Ito 8色順に固定し、accent(clay)は「強調1系列」専用に予約、残りはグレースケール。テーマトークンに `series_muted`(グレー)と `series_accent`(1色)を追加。8系列超でエラー。出典:Wong / six minutes。`builder.py` + `theme.py`。

**M4. `table` 型をミニマル罫線に変更(原則12)**
縦罫廃止・ヘッダ下と最下行のみ hairline(既存 `HAIRLINE_W` 流用)・中間行は罫なし。数値列は自動右揃え(混在は小数点揃え)、等幅数字(tabular-nums相当フォント指定)。行数 > 7 で警告「変数を1–3個に絞るか分割」(査読PMC由来=学会で説得力大)。出典:CSE / PMC。

**M5. 新型 `big-number`(巨大数字単独スライド)(原則3/第2章)**
画面中央に超大型(60–120pt、`SZ_METRIC=44pt` を超える専用スケール)+下にsmall-caps説明1行+細いアクセント線。既存 kpi は複数指標前提なので単一データ点を分離。背景反転(ink背景×cream文字)オプション付与。出典:piktochart「one data point per slide」/ SlideGenius。

**M6. 新型 `statement`/`big-statement`(全画面1メッセージ)(原則1/14)**
地色全面+中央に1行の巨大テキスト(60–100pt、font_scale連動)のみ。`quote`/`takeaway` と差別化し章転換・ピッチの転換点に。純白回避(地色維持)を強制。出典:Duarte コントラスト=単一焦点 / 高橋メソッド。

**M7. `dark` モディファイア(背景反転)(第2章)**
`_class: dark` で地色 ink `#141413`・文字 cream・accent は clay のまま。章転換・強調演出に。反転側でもコントラスト再検査(H3と連動)。出典:SlideGenius「dark gradient + white text」。

**M8. データ系全型に `source:` フロントマター + 警告(原則・脚注規律)**
`figure/table/result/kpi/multi-result/funnel/stack/chart` に `source:` を追加し、未指定なら警告。配置はスライド下端・small-caps・本文より2pt小、「modified from …」自動整形。ページ番号も全型デフォルトON。出典:Slideworks / Sheridan。

**M9. `--storyline`(横の論理ビュー/ゴーストデッキ出力)(原則4)**
全スライドのH2(アクションタイトル)だけを抽出し、アウトライン/プレーンテキストで出力。「タイトルだけ読んで物語が通るか」を著者が検証。Web UIなら左ペインにタイトル一覧。出典:Slideworks(横の論理・ゴーストデッキ)。

### Low(将来拡張・効果限定 or 要実測)

**L1. 円グラフ `pie` 入力時の警告 + 棒への代替提案(原則12)**
要素数 > 5 または僅差ランキング検出で「円は僅差比較に不向き(棒推奨)。割合提示目的なら維持可」。全面禁止せずタスク依存の助言に(SAGE知見)。

**L2. 色覚シミュレーション出力 `--simulate deuteranopia|protanopia|tritanopia|grayscale`(原則13)**
確認用PNGを生成し系列が判別不能なら警告。matplotlib PNG fallback パイプラインに相乗り。出典:JFLY CUD。

**L3. デッキ尺メーター `--target-minutes 20`(原則3)**
ビルド時に総枚数を「≈N分トーク相当(1枚≈1分)」と表示、目標と大きく乖離で警告。研究発表テンプレで `agenda` を既定に入れない設計に。出典:PLOS Rule 2。

**L4. 構成スキャフォールド `--scaffold conference|business`(原則14/第2章)**
- conference:`motivating-example(例から入る)→ rq → method → result/multi-result → takeaway`
- business:`scqa → executive-summary(結論+3論拠)→ … → takeaway`
新型 `motivating-example` / `scqa` / `executive-summary` を併設。出典:Peyton Jones / think-cell(ピラミッド/SCQA)。

**L5. 「テキストのみスライド」検出(原則10)**
figure/diagram/table/equation/chart のいずれも無く本文だけのスライドで警告。`takeaway`/`quote` は型側でオプトアウト可。出典:PLOS Rule 6。

**L6. トランジション無効をデフォルトに(原則9)**
出力pptxのトランジションは無効を既定、設定で限定有効化(過剰アニメ抑止をツール規律として明文化)。出典:Garr Reynolds。

**L7. 冗長符号化レンダリング `series-legend`(原則13)**
系列を色だけでなく形(○△□)/線種(実線/点線)/ハッチングで重複コード化。`pros-cons` は緑/赤でなく ✓/✗ アイコン+青/橙(または緑/マゼンタ)に。出典:JFLY CUD。

---

## 4. 参考実例・ギャラリー(実在確認済みのみ)

**原則・規範(一次/精読確認済)**
- PLOS Comp Biol「Ten simple rules for effective presentation slides」— https://journals.plos.org/ploscompbiol/article?id=10.1371/journal.pcbi.1009554 — 査読付き定量規範(6要素/1分1枚/見出しに結論)。学会用途の根拠として最良。
- Simon Peyton Jones「How to give a good research talk」— https://www.microsoft.com/en-us/research/wp-content/uploads/2016/08/giving-a-talk.pdf — CS研究トークの古典。例の重視・目次禁止・時間厳守。
- Garr Reynolds「Presentation Design Tips」— https://www.garrreynolds.com/design-tips — Presentation Zen 著者の一次原則(余白/SNR/コントラスト/画像優位性)。
- TED「Create + prepare slides」— https://www.ted.com/pages/create-prepare-slides — 公式スライド規定(1論点/本文≤6行)。
- Butterick「Practical Typography」— https://practicaltypography.com/line-length.html — 行長・サイズ・余白を数値で示す実務家の権威。
- W3C WCAG 2.1 / 1.4.3 — https://www.w3.org/WAI/WCAG21/Understanding/contrast-minimum.html — コントラスト比の規範的一次ソース。
- WebAIM Contrast Checker — https://webaim.org/resources/contrastchecker/ — トークン色のAA/AAA合否判定。
- Baymard「Line-length readability」— https://baymard.com/blog/line-length-readability — 行長50–75字・line-height 1.5em の実証。
- BrightCarbon「Presentation font size」/「Advanced grids and guides」— https://www.brightcarbon.com/blog/presentation-font-size/ / https://www.brightcarbon.com/blog/advanced-powerpoint-grids-guides/ — 用途別フォント下限・12カラム/セーフ余白。
- spec.fm「Type Scale」— https://spec.fm/specifics/type-scale — モジュラースケール比率の早見。

**データ可視化・色覚**
- Bang Wong「Color blindness」Nature Methods 2011 — https://www.nature.com/articles/nmeth.1618 — 色覚バリアフリー8色パレットの規範。
- JFLY / Color Universal Design — https://jfly.uni-koeln.de/color/ — 配色・凡例・レーザーポインタまでの実装策。
- Okabe-Ito hex リファレンス — https://conceptviz.app/blog/okabe-ito-palette-hex-codes-complete-reference — 8色hex+グレースケール挙動。
- CSE「Best Practices in Table Design」— https://www.csescienceeditor.org/article/best-practices-in-table-design/ — ミニマル表設計の決定版。
- Hill 2025「Are pie charts evil?」(SAGE) — https://journals.sagepub.com/doi/10.1177/14738716241259432 — 円vs棒のタスク別最新実証。
- Six Minutes「Slide Charts: 20 Guidelines」— https://sixminutes.dlugan.com/slide-charts/ — 発表者向けチャート20則。

**コンサル/キーノート/プロダクト**
- Slideworks「Action titles」/「How McKinsey make presentations」— https://slideworks.io/resources/how-to-write-action-titles-like-mckinsey / https://slideworks.io/resources/how-mckinsey-consultants-make-presentations — アクションタイトル/横の論理/ゴーストデッキ。
- think-cell「Pyramid Principle」— https://www.think-cell.com/en/resources/content-hub/using-the-pyramid-principle-to-build-better-powerpoint-presentations — SCQA→ピラミッド→MECE。
- Deckary「Consulting Slide Standards」— https://deckary.com/blog/consulting-slide-standards — 15語/2フォント/3-4色/60秒の定量一覧。
- Duarte「Slide:ology」/「Ultimate Guide to Contrast」— https://www.duarte.com/resources/books/slideology/ / https://www.duarte.com/blog/ultimate-guide-to-contrast/ — 30語/5論点/75語ルール、コントラスト=単一焦点。
- ideas.ted.com「6 dos and don'ts」— https://ideas.ted.com/6-dos-and-donts-for-next-level-slides-from-a-ted-presentation-expert/ — 色/~30pt/空白スライド戦略。
- PMC「Two Minutes More!」— https://pmc.ncbi.nlm.nih.gov/articles/PMC9896115/ — 表5–7行・導入1–2枚・7–10分の学術一次規範。
- MIT Comm Lab「Redesigning Existing Figures for Slides」— https://mitcommlab.mit.edu/cee/2024/01/15/redesigning-existing-figures-for-slides/ — 論文図→スライド図の before/after。

**ギャラリー(実例集・Marpエコシステム)**
- Deck.gallery — https://www.deck.gallery/ — Nike/Disney/IBM等を design quality でキュレーション。型設計の対照表。
- Marp Community Themes — https://rnd195.github.io/marp-community-themes/ — 11テーマをLight/Darkで分類(直接の競合比較対象)。
- awesome-marp — https://github.com/marp-team/awesome-marp — Marp公式の厳選(Beamer風〜モダン)。
- IEEE ProComm「Maximize Signal to Noise Ratio」— https://procomm.ieee.org/design-your-slides-to-maximize-signal-to-noise-ratio/ — 工学コミュニティ向けSNR論。

(検索実在確認のみ・本文未精読:Slidebean ピッチデック集 https://slidebean.com/blog/best-startup-pitch-decks-of-all-time、Figma pitch deck examples https://www.figma.com/resource-library/pitch-deck-examples/、Pitch templates https://pitch.com/templates/collections/Pitch-deck、Design Shack Keynote templates https://designshack.net/articles/inspiration/best-keynote-templates/、Information is Beautiful https://informationisbeautiful.net/。Gamma設計ガイドは403で本文未取得。)

---

## 5. 「やってはいけない」アンチパターン集

各項目に対応する違反原則と出典を併記。

1. **見出しを名詞ラベルにする**(「結果」「Method」「Results」)。結論文(主張)にしないと自己完結性・横の論理が崩れる。→原則4。PLOS Rule 3 / Slideworks。
2. **論文の図をそのまま貼る**。情報密度が高すぎ聴衆が読解に没入し話を聞かなくなる。1スライド1パネルに再描画せよ。→原則10。MIT Comm Lab / Paperpile。
3. **1スライドに7要素以上・本文を文章で詰め込む**(75語超は「文書」)。読む/聞くで注意が分散。→原則2。PLOS Rule 7 / Duarte。
4. **clay `#d97757` を cream 上の本文色に使う**(実測≈3.0:1でAA未達)。本文は ink、clay は線・大見出し・大数字専用。→原則8。WCAG 1.4.3。
5. **赤×緑を隣接させる/色だけで系列を区別する**。男性8%が判別困難。Okabe-Ito+形/線種/直接ラベルを併用。→原則13。Wong / JFLY。
6. **円グラフで僅差のランキングを見せる/要素過多の円**。長さ(棒)の方が約1.96倍正確で速い。→原則12。Hill 2025 / Cleveland。
7. **表に縦罫を全部引く・数値を左揃え**。縦罫は余白で代替、数値は右揃え/小数点揃え・等幅。→原則12。CSE。
8. **目次スライドで始める/「introduction」「conclusion」など自明項目を立てる**。貴重な冒頭1分を浪費。動機づけの例から入れ。→原則14。Peyton Jones §2。
9. **抽象論・記号(squiggles)だけで例を省く(The Awful Trap)**。聴衆を置き去りにする最頻の失敗。→原則7(例の重視)。Peyton Jones §2.1。
10. **真っ白スライドや過剰トランジション/アニメーション**。純白は機材故障と誤認、アニメはノイズ。地色維持・トランジションは原則無効。→原則9/14。ideas.ted / Garr Reynolds。
11. **データ図・表に出典脚注が無い**。学会・コンサル両用途で必須。下端small-caps+末尾参照スライド。→脚注規律。Slideworks / Sheridan。
12. **1スライド3種超のフォントサイズ/3書体以上/全大文字・下線・イタリック多用**。階層が崩れ可読性低下。サイズ≤3・書体≤2、強調は太字。→原則5。brightcarbon / PLOS Rule 7。
13. **大会場で本文18pt以下**。後方席から読めない。会場距離に応じ24–30pt+へ。→原則5。thepopp / brightcarbon。
14. **1スライドに2つ以上のメッセージ(タイトルに「and」)**。分割サイン。1スライド1チャート1メッセージ。→原則1。Deckary。

---

注記(未確認・未精読の明示):余白の権威ある具体%、Duarte「30語/5論点/75語」の一次ページ逐語、Carmine Gallo「19語/10-12枚」の本書本文、Cleveland 1.96倍の原論文本文、Okabe-Ito個別hexのNature本文(ペイウォール内)は一次精読ができず、複数の二次要約/検索スニペットで一致を確認した範囲で採用。Gamma設計ガイドは403で本文未取得。これら以外の主要定量値(WCAG 4.5:1/3:1、6要素/1分1枚、表5–7行、行長45–75字、行間1.2–1.5、ジャンプ率)は一次/精読確認済み。コード照合(`layout.py` トークン、`claude.css` 配色)は本タスク内で直接確認済み。