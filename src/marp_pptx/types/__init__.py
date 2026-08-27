"""Slide type registry — semantic metadata for all slide types."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class SlideTypeInfo:
    name: str
    css_class: str
    category: str
    geometry: str
    meaning: str
    use_when: str
    template_file: str


TYPE_REGISTRY: list[SlideTypeInfo] = [
    # ── Structure ──
    SlideTypeInfo("cols-2", "cols-2", "structure", "□ □", "並列・対比", "2つを同じ重みで比較するとき", "03-cols-2.md"),
    SlideTypeInfo("cols-3", "cols-3", "structure", "□ □ □", "分類・カテゴリ", "3つの側面を示すとき", "04-cols-3.md"),
    SlideTypeInfo("sandwich", "sandwich", "structure", "─ □□ ─", "概要→詳細→結論", "フレーミングが必要なとき", "05-sandwich.md"),
    SlideTypeInfo("sandwich-3col", "sandwich-3col", "structure", "─ □□□ ─", "概要→3並列→結論", "共通設定＋3条件＋考察を1枚にまとめるとき", "06-sandwich-3col.md"),
    SlideTypeInfo("split-text", "split-text", "structure", "□│□", "二面性・補完", "左右で補完的な内容を示すとき", "45-split-text.md"),
    SlideTypeInfo("card-grid", "card-grid", "structure", "□□ □□", "均質な要素の一覧", "同種のアイテムを並べるとき", "44-card-grid.md"),

    # ── Temporal ──
    SlideTypeInfo("timeline-h", "timeline-h", "temporal", "●─●─●─●", "時系列・経過", "時間の流れを示すとき", "15-timeline-horizontal.md"),
    SlideTypeInfo("timeline-v", "timeline", "temporal", "●│●│●", "手順・段階", "縦の時系列を示すとき", "14-timeline-vertical.md"),
    SlideTypeInfo("steps", "steps", "temporal", "①→②→③", "プロセス・手順", "順にやる手順を示すとき", "29-steps.md"),
    SlideTypeInfo("before-after", "before-after", "temporal", "□ → □", "変化・改善", "ビフォーアフターを示すとき", "41-before-after.md"),
    SlideTypeInfo("history", "history", "temporal", "年│出来事", "沿革・文脈", "歴史的経緯を示すとき", "31-history.md"),

    # ── Convergence ──
    SlideTypeInfo("funnel", "funnel", "convergence", "▽", "絞り込み・選別", "多→少の過程を見せるとき", "42-funnel.md"),
    SlideTypeInfo("stack", "stack", "convergence", "□ □ □ 積層", "積み上げ・累積", "レイヤー構造を示すとき", "43-stack.md"),
    SlideTypeInfo("overview", "overview", "convergence", "大□＋小□群", "全体像と部分", "全体→詳細の構造を示すとき", "27-overview.md"),
    SlideTypeInfo("highlight", "highlight", "convergence", "███ ■ ███", "強調・焦点", "1つだけ際立たせるとき", "38-highlight.md"),

    # ── Evaluation ──
    SlideTypeInfo("pros-cons", "pros-cons", "evaluation", "＋│−", "賛否・長短", "判断材料を示すとき", "34-pros-cons.md"),
    SlideTypeInfo("zone-compare", "zone-compare", "evaluation", "□ vs □", "比較評価", "2つを評価比較するとき", "19-zone-compare.md"),
    SlideTypeInfo("zone-matrix", "zone-matrix", "evaluation", "□□ □□ (2x2)", "二軸評価", "2軸で分類するとき", "20-zone-matrix.md"),
    SlideTypeInfo("kpi", "kpi", "evaluation", "数字 数字 数字", "定量評価", "KPI・数値を強調するとき", "33-kpi.md"),
    SlideTypeInfo("result", "result", "evaluation", "結論＋根拠", "成果報告", "実験結果を報告するとき", "28-result.md"),
    SlideTypeInfo("result-dual", "result-dual", "evaluation", "結果□│結果□", "二つの成果を並列", "2つの結果を比較するとき", "24-result-dual.md"),
    SlideTypeInfo("multi-result", "multi-result", "evaluation", "結果□□□", "複数成果の一覧", "複数の結果を一覧するとき", "47-multi-result.md"),
    SlideTypeInfo("big-number", "big-number", "evaluation", "巨大数字", "単一指標の強調", "ひとつの数字を主役にするとき", "51-big-number.md"),
    SlideTypeInfo("chart", "chart", "evaluation", "棒/折れ線グラフ", "データの可視化", "数値データをグラフで見せるとき(編集可)", "52-chart.md"),

    # ── Knowledge ──
    SlideTypeInfo("definition", "definition", "knowledge", "用語：説明", "概念の定義", "用語を定義するとき", "35-definition.md"),
    SlideTypeInfo("equation", "equation", "knowledge", "$式$ 中央", "数理的真理", "数式が主役のとき", "07-equation.md"),
    SlideTypeInfo("equations", "equations", "knowledge", "$式$ $式$ $式$", "式の体系", "連立方程式・最適化問題のとき", "17-equations-opt.md"),
    SlideTypeInfo("equation-annotated", "equation-annotated", "knowledge", "$式$ ＋ 凡例", "数式＋記号注釈", "数式の各記号を注釈付きで解説するとき", "08-equation-annotated.md"),
    SlideTypeInfo("equation-highlight", "equation-highlight", "knowledge", "$式$ 領域強調", "数式の着目部を強調", "数式の特定項を色で強調して読み解くとき", "09-equation-highlight.md"),
    SlideTypeInfo("diagram", "diagram", "knowledge", "図＋説明", "構造の可視化", "図で構造を説明するとき", "36-diagram.md"),
    SlideTypeInfo("annotation", "annotation", "knowledge", "図＋注釈", "詳細解説", "図に注釈を付けるとき", "40-annotation.md"),
    SlideTypeInfo("code", "code", "knowledge", "コードブロック", "実装・手続き", "コードを見せるとき", "46-code.md"),
    SlideTypeInfo("blocks", "blocks", "knowledge", "▬題│本文 ×N", "定理・定義ブロック", "beamer 流の定理環境（theorem/example/alert）を並べるとき", "53-blocks.md"),
    SlideTypeInfo("sections", "sections", "knowledge", "■リード│本文 ×N", "高密度トピック積層", "1枚に2〜4トピックを「色付きリード行＋本文」の帯で積むとき（公聴会流の高密度）", "54-sections.md"),
    SlideTypeInfo("flow", "flow", "flow", "□→□→□ +loop", "ブロック図・ループ図", "mermaid flowchart 記法から編集可能なブロック図/機器構成図/反復ループ図を描くとき", "55-flow.md"),
    SlideTypeInfo("paper", "paper", "knowledge", "▤書誌+理由+要点", "文献報告の書誌カード", "輪読・文献調査で紹介論文の書誌情報・選定理由・要点を1枚にするとき", "56-paper.md"),
    SlideTypeInfo("split-panel", "split-panel", "structure", "▮左色面│本文", "ハーフブリード色面パネル", "画面端まで塗った色面に主張を白抜きし、右に本文を置くとき（キーノート級の1枚）", "57-split-panel.md"),
    SlideTypeInfo("graphical-abstract", "graphical-abstract", "structure", "□→□→▪ 一枚絵", "グラフィカルアブストラクト", "表紙直後に課題→手法→成果を1枚の図で示すとき（研究発表の定番）", "58-graphical-abstract.md"),
    SlideTypeInfo("figure-full", "figure-full", "structure", "▣ 最大画角の図", "論文図の全面表示", "論文の特徴的な図を余白0.25inまで最大サイズで見せるとき（図が主役の1枚）", "59-figure-full.md"),
    SlideTypeInfo("figure-story", "figure-story", "structure", "▤図│解説＋結論帯", "図の完全解剖", "図を左右どちらかに寄せ、リード文（h2）・横の解説・下のまとめ帯で1枚を完結させるとき", "60-figure-story.md"),

    # ── Flow ──
    SlideTypeInfo("zone-flow", "zone-flow", "flow", "□→□→□", "フロー・因果", "原因→結果の流れを示すとき", "18-zone-flow.md"),
    SlideTypeInfo("zone-process", "zone-process", "flow", "□→□→□ (詳細)", "プロセス＋詳細", "詳細付きプロセスを示すとき", "21-zone-process.md"),
    SlideTypeInfo("agenda", "agenda", "flow", "1. 2. 3.", "予定・構成", "発表の構成を示すとき", "22-agenda.md"),
    SlideTypeInfo("checklist", "checklist", "flow", "☑ ☑ ☐", "進捗・完了状態", "タスクの状態を示すとき", "39-checklist.md"),

    # ── Narrative ──
    SlideTypeInfo("quote", "quote", "narrative", "「　」", "権威・声", "引用を示すとき", "30-quote.md"),
    SlideTypeInfo("profile", "profile", "narrative", "写真＋経歴", "人物紹介", "人物を紹介するとき", "49-profile.md"),
    SlideTypeInfo("takeaway", "takeaway", "narrative", "★ メッセージ", "持ち帰ってほしい1つ", "キーメッセージを伝えるとき", "48-takeaway.md"),
    SlideTypeInfo("statement", "statement", "narrative", "全画面 1文", "断言・転換", "1メッセージを大きく言い切るとき(dark可)", "50-statement.md"),
    SlideTypeInfo("panorama", "panorama", "narrative", "横幅画像", "インパクト・没入", "大きな画像で印象づけるとき", "32-panorama.md"),
    SlideTypeInfo("gallery-img", "gallery-img", "narrative", "画像群", "ビジュアル一覧", "複数画像を並べるとき", "37-gallery-img.md"),
    SlideTypeInfo("figure", "figure", "narrative", "画像＋キャプション", "図の提示", "図を中心に見せるとき", "10-figure.md"),
    SlideTypeInfo("figure-cols", "figure-cols", "narrative", "図 │ 解説", "図＋考察の対比", "図と解説（観察・含意）を左右に並べるとき", "11-figure-cols.md"),

    # ── Meta ──
    SlideTypeInfo("title", "title", "meta", "大タイトル", "始まり", "プレゼンの冒頭", "01-title.md"),
    SlideTypeInfo("divider", "divider", "meta", "セクション区切り", "転換", "章の区切り", "02-divider.md"),
    SlideTypeInfo("summary", "summary", "meta", "まとめリスト", "まとめ", "内容を要約するとき", "25-summary.md"),
    SlideTypeInfo("rq", "rq", "meta", "中央に問い", "問いの提示", "研究質問を示すとき", "23-rq.md"),
    SlideTypeInfo("references", "references", "meta", "文献リスト", "学術的裏付け", "参考文献を示すとき", "13-references.md"),
    SlideTypeInfo("appendix", "appendix", "meta", "補足", "補足資料", "追加情報を示すとき", "26-appendix.md"),
    SlideTypeInfo("end", "end", "meta", "Thank You", "終わり", "プレゼンの終了", "16-end.md"),
    SlideTypeInfo("table-slide", "table-slide", "meta", "表", "データ表示", "表形式でデータを示すとき", "12-table.md"),
]


def get_type_info(css_class: str) -> SlideTypeInfo | None:
    for t in TYPE_REGISTRY:
        if t.css_class == css_class:
            return t
    return None


CATEGORIES = {
    "structure": "構造",
    "temporal": "時間",
    "convergence": "収束・拡散",
    "evaluation": "評価・判断",
    "knowledge": "知識・定義",
    "flow": "流れ・構造",
    "narrative": "ナラティブ",
    "meta": "メタ",
}
