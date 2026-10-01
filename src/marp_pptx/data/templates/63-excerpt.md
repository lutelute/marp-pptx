---
marp: true
theme: academic
paginate: true
---

<!-- _class: excerpt -->

# 原典を読む｜Transformer の出発点

## 問題設定｜再帰も畳み込みも使わずに系列を変換する（Vaswani et al., 2017）

<div class="ex-item">
<span class="ex-quote">"The dominant sequence transduction models are based on complex recurrent or convolutional neural networks that include an encoder and a decoder."</span>
<span class="ex-cite">Abstract</span>
<span class="ex-read">当時の主流は==再帰か畳み込み==を前提にしていた</span>
</div>

<div class="ex-item">
<span class="ex-quote">"We propose a new simple network architecture, the Transformer, based solely on attention mechanisms, dispensing with recurrence and convolutions entirely."</span>
<span class="ex-cite">Abstract</span>
<span class="ex-read">注意機構==だけ==で組み、再帰と畳み込みを完全に外した</span>
</div>

<div class="ex-verdict">**利点**　再帰が無いので、系列の全位置を並列に計算できる<br>**課題**　注意の計算量は系列長の 2 乗 — ==長い系列ほど重い==</div>

<div class="ex-source">A. Vaswani, N. Shazeer, N. Parmar, J. Uszkoreit, L. Jones, A. N. Gomez, Ł. Kaiser and I. Polosukhin, "Attention is all you need," NeurIPS, 2017.</div>
