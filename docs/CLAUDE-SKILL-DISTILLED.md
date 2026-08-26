# Claude スライド機能の蒸留 → marp-pptx への落とし込み

2026-08-26。Anthropic 公式の Claude スライド生成ノウハウ（`document-skills:pptx` スキルの
Design Ideas / Avoid リスト / QA 手順、`theme-factory` のテーマ集）を読み解き、
marp-pptx の**決定論エンジン側**に移植した記録。何をどこに落としたか・意図的に
採用しなかったものは何かを残す。広域の設計リサーチは
[SLIDE-DESIGN-RESEARCH.md](SLIDE-DESIGN-RESEARCH.md)、実装ロードマップは
[FEATURE-DESIGN.md](FEATURE-DESIGN.md) を参照。

## 採用マッピング

| 蒸留元（pptx スキルの原則） | 落とし込み先 |
|---|---|
| Dark/light **sandwich**（表紙と結びをダークに、本文は明るく） | `config-claude.yaml` を `title_bg: dark` / `end_bg: dark` に変更。新パレット4種も同構成 |
| カードの区別は「地色の差＋シャドウ」で、エッジの色帯でやらない | `ThemeLayout.card_shadow` を実装（従来は死にトークン）。`_soft_shadow()` が 15% 黒・7pt ぼかしの outerShdw を注入。claude と新4テーマで有効 |
| タイトルはサイズで支配する（36-44pt、本文と強い対比） | `ThemeLayout.title_scale` 新設。claude 系は 1.13（H1 30→34pt 相当、ヒーロー 40→45pt）。長い見出しは従来どおり実測ベースで自動縮小するので、短い断言だけが大きくなる |
| チャートは**デッキのパレットで**塗る（既定色のままにしない） | `_chart_colors()`: accent 主導のテーマ由来ランプ（accent → primary → accent の淡色 → secondary）。輝度比 1.3 未満のペアが出るテーマと 5 系列以上は Okabe-Ito にフォールバック（実測ゲート） |
| ダーク面ではアクセントを活かす（低コントラストの飾りを置かない） | `_hero_accent(min_ratio)`: PRIMARY 地の上でテーマ accent の WCAG 比を実測し、罫線は 2.5、キッカー文字は 4.5 を下回ったら白へ。claude はクレイがインク地で生きる（5.9:1） |
| 大胆で題材に紐づくパレット（"generic blue に逃げない"） | pptx スキルのパレット表から 4 種を移植: `midnight`（経営）/ `terracotta`（教育）/ `teal`(医療・公共) / `cherry`（キーノート）。全て accent_text は白地 4.5:1 以上を実測して選色 |
| パレットは題材で選ぶ・全スライド視覚アンカー・同レイアウト連続禁止・大数字は callout | `skills/marp-pptx/SKILL.md` の「デザイン原則」節（オーサリング時の規範として） |

## すでに一致していた（今回の変更なし）

- **「タイトル下のアクセント線は AI スライドの証」** — `_add_title` は下線を既定で
  描かず、docstring に同じ理由を明記済み。装飾はテーマの明示オプトインのみ
- **QA ループ（レンダ→目視→修正）** — `doctor`（実測 overflow/overlap/contrast/font/deck）
  と MCP `preview_png` が既に決定論で担う
- 本文左揃え・タイトルのみ中央 / 0.5in マージン / ブロック間 0.3in 前後 — layout.py の
  トークンが既に満たす

## 意図的に採用しなかったもの

- **「cream 背景を既定にしない」** — claude テーマの cream は Anthropic ブランドの核で
  ユーザー決定済み。新規4テーマの本文背景は白 `#ffffff` にしてこの原則側に置いた
- **Okabe-Ito の全廃** — 色覚多様性の安全性は学術用途の要件。テーマランプが
  縮退するテーマと5系列以上では今も既定
- **アイコン（react-icons レンダ）** — MD パイプラインに画像生成を持ち込むと
  「編集可能 PPTX」の決定論が崩れる。番号サークル（steps/agenda）が既にその役
- **フォント安全リスト（Arial/Calibri 系）** — 生成環境がユーザーの Mac ローカルで
  あり、doctor の `font` 検査が実測で代替している

## 検証

- `demo/hosting-capacity-v3.md` を claude / midnight で実レンダし目視（sandwich・
  シャドウ・チャート色・タイトルスケールを確認）
- パレット4種の accent_text / white、white / primary、accent / primary は
  すべて WCAG 実測値で選定（このリポの方針: 推定せず測る）
