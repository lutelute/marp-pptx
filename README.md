# Marp Academic Slide Templates

学会発表用の Marp スライドテンプレート集。

## Quick Start

```bash
# プレビュー
npx @marp-team/marp-cli --theme themes/academic.css -p example.md

# PDF出力
npx @marp-team/marp-cli --theme themes/academic.css --pdf --allow-local-files example.md

# PPTX出力
npx @marp-team/marp-cli --theme themes/academic.css --pptx --allow-local-files example.md
```

## Structure

```
themes/academic.css    ← スライドマスター（テーマCSS）
templates/             ← スライド種類別テンプレート（個別ファイル）
example.md             ← 全テンプレートを使ったサンプル発表
```

## Slide Classes

| Class | Description |
|-------|-------------|
| `title` | タイトルスライド |
| `divider` | セクション区切り |
| `cols-2` / `cols-2-wide-l` / `cols-2-wide-r` | 2カラム |
| `cols-3` | 3カラム |
| `sandwich` | 上下全幅 + 中央マルチカラム |
| `equation` | 数式中央配置 + 変数説明 |
| `figure` | 図キャプション + 解説 |
| `table-slide` | 表スタイリング |
| `references` | 参考文献リスト |
| `timeline` / `timeline-h` | 歴史フロー（縦/横） |
| `end` | 終了スライド |

## Utility Classes

| Class | Description |
|-------|-------------|
| `.box` | グレー背景ボックス |
| `.box-accent` | 左ボーダー赤のボックス |
| `.box-primary` | 左ボーダー青のボックス |
| `.eq-highlight` / `.eq-highlight-b` | 数式ハイライト（黄/青） |
| `.footnote` | スライド下部の脚注 |
| `.small` `.muted` `.bold` `.center` | テキストユーティリティ |

## Preview

### Title
![Title](docs/slide.001.png)

### Section Divider
![Divider](docs/slide.002.png)

### Content + Box
![Content](docs/slide.003.png)

### Horizontal Timeline
![Timeline](docs/slide.004.png)

### Equation (Highlight)
![Equation Highlight](docs/slide.006.png)

### Equation (Underbrace)
![Equation Underbrace](docs/slide.007.png)

### Sandwich Layout (3-col)
![Sandwich](docs/slide.008.png)

### Table
![Table](docs/slide.010.png)

### 2-Column with Figure
![2-col Figure](docs/slide.011.png)

### 3-Column Comparison
![3-col](docs/slide.012.png)

### Summary (Sandwich 2-col)
![Summary](docs/slide.014.png)

### References
![References](docs/slide.015.png)

### End
![End](docs/slide.016.png)
