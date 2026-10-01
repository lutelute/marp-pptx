---
marp: true
theme: academic
paginate: true
---

<!-- _class: board -->
<!-- _kicker: 基礎 ｜ データ形式 -->

# GeoJSON は「地理を JSON で書く」地理フォーマット

<div class="defn">

**GeoJSON（Geographic JSON・RFC 7946）**
JSON の書き方で「点・線・面」の地物と属性を表す。座標は [経度, 緯度] の順で並べる。

</div>

<div class="row">
<div class="col">

#### 3 つのプリミティブ

### :point: Point ｜ 点
変電所・発電所などの地点

### :line: LineString ｜ 線
送電線の経路（頂点の並び）

### :polygon: Polygon ｜ 面
供給エリア・発電所の敷地

<div class="callout">→ 系統の構成要素は すべて 点・線・面 に集約<br>==地理データのまま 系統モデルへ==</div>

</div>
<div class="col">

#### 例：変電所を GeoJSON で書くと

```geojson
{
  "type": "Feature",
  "geometry": {
    "type": "Point",
    "coordinates": [136.22, 36.06]
  },
  "properties": {
    "name": "福井変電所",
    "voltage": 500
  }
}
```

</div>
</div>

<!-- source: 出典：RFC 7946 The GeoJSON Format (IETF, 2016) -->
