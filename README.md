# jma-code-dictionary
防災情報XMLのコード表など、気象庁が公開しているデータをCSVに変換して配置しているリポジトリ  
毎日自動で確認しているため、ある程度であれば自動で変更が追従されます

## ディレクトリ構成

出典ごとにディレクトリを分けています。

| ディレクトリ | 内容 | 出典 |
|---|---|---|
| `csv/` | 防災情報XMLのコード表（個別コード表・コード管理表一式） | [気象庁防災情報XMLフォーマット 技術資料](https://xml.kishou.go.jp/tec_material.html) |
| `csv/web-json/` | 気象庁ホームページの「防災情報」で使用されている JSON | [気象庁 防災情報](https://www.jma.go.jp/bosai/) |
| `csv/bulletin/` | 地震月報（カタログ編）の資料 | [地震月報（カタログ編） 検測値データ](https://www.data.jma.go.jp/eqev/data/bulletin/deck.html) |
| `csv/multilingual/` | 気象情報等で用いる用語の多言語辞書 | [多言語辞書データ](https://www.data.jma.go.jp/developer/multilingual.html) |

`csv/web-json/` の元になっている JSON は仕様が公開されていないため、予告なく構造や URL が変わる可能性があります。

## 更新の確認

取得元ごとの `Last-Modified` を `.cache/lastModified.json` に記録し、前回から更新されたものだけを取得して CSV を作り直しています。  
生成処理を変更したときなど、強制的に作り直したい場合は、このファイルを削除してください。  
項目ごとに削除することもできますが、コード表（`codeManagementTable`・`individualCodeTables`）は技術資料のページ（`techMaterialPage`）が更新されたときだけ確認するため、あわせて `techMaterialPage` も削除してください。

## CSV の形式

- 文字コードは UTF-8（`csv/AmedasRainPoint.csv` のみ BOM 付き）、改行は LF です。
- ヘッダー行はありません。
- カンマを含む値は `"` で囲み、値の中の `"` は `\"`、改行は `\n` として出力しています。
- 値は元データの表記のまま出力しています（単位の換算などは行っていません）。

## コード表以外の CSV の列

### `csv/web-json/Volcano.csv` 火山

出典: https://www.jma.go.jp/bosai/volcano/const/volcano_list.json

| 列 | 内容 | 元の項目 |
|---|---|---|
| 1 | 火山コード（`csv/PointVolcano.csv` と共通） | `code` |
| 2 | 火山名 | `name_jp` |
| 3 | 火山名（英語） | `name_en` |
| 4 | 緯度（度） | `latlon[0]` |
| 5 | 経度（度） | `latlon[1]` |
| 6 | `true` または空（噴火警戒レベルの運用に関する値と思われますが、意味は公開されていません） | `levelOperation` |

コード 900～902（全国の活火山・その他の活火山・新たな活火山）には位置がありません。

### `csv/web-json/TideLevelStation.csv` 潮位観測点

出典: https://www.jma.go.jp/bosai/tidelevel/const/tide_area.json

| 列 | 内容 | 元の項目 |
|---|---|---|
| 1 | 観測点コード（防災情報XMLの津波観測点コードとは別の体系です） | `code` |
| 2 | 観測点名 | `name` |
| 3 | 緯度（度） | `lat` |
| 4 | 経度（度） | `lon` |
| 5 | 所在地 | `addr` |
| 6 | 観測機関の略号（`jma`, `phb` など） | `type` |
| 7 | 観測機関名 | `typeName` |
| 8 | 潮位の基準 | `reference` |
| 9 | 観測方式 | `detail.type` |
| 10 | 標高への換算に関する値と思われますが、意味は公開されていません | `detail.height1` |
| 11 | 同上 | `detail.height2` |
| 12 | 同上 | `detail.height3` |
| 13 | 備考 | `detail.note` |
| 14 | 過去最高潮位（cm） | `max.level` |
| 15 | 過去最高潮位の日時（yyyymmddhhmm） | `max.datetime` |
| 16 | 過去最高潮位の統計期間 | `max.period` |
| 17 | 過去最高潮位の要因 | `max.description` |
| 18 | `true` または空 | `missing` |
| 19 | 二次細分区域コード（`csv/AreaInformationCity.csv` と共通） | `parents.class20` |
| 20 | 地域コード（意味は公開されていません） | `parents.class30` |

気象庁以外の機関（港湾局・海上保安庁・地方公共団体など）が管理する観測点を含みます。

### `csv/web-json/WindProfilerStation.csv` ウィンドプロファイラ観測局

出典: https://www.jma.go.jp/bosai/windprofiler/const/station.json

| 列 | 内容 | 元の項目 |
|---|---|---|
| 1 | 地点番号 | （キー） |
| 2 | 地点名 | `jpname` |
| 3 | 地点名（ローマ字） | `enname` |
| 4 | 緯度（度） | `lat` |
| 5 | 経度（度） | `lon` |
| 6 | 高度 | `altitude` |

### `csv/bulletin/SeismicStation.csv` 地震観測地点

出典: https://www.data.jma.go.jp/eqev/data/bulletin/data/deck/stations.zip  
フォーマット: [地震観測点ファイルフォーマット](https://www.data.jma.go.jp/eqev/data/bulletin/data/format/stnfmt_j.html)（列の順序はこのフォーマットのとおりです）

| 列 | 内容 |
|---|---|
| 1 | 観測点コード |
| 2 | 経度（度） |
| 3 | 経度（分） |
| 4 | 緯度（度） |
| 5 | 緯度（分） |
| 6 | 観測点番号（部内用） |

分の列は元のファイルでは小数点を省略した形式（F4.2）のため、小数点を補って出力しています（例: `1280` → `12.80`）。

### `csv/multilingual/MultilingualDictionary.csv` 多言語辞書

出典: https://www.data.jma.go.jp/developer/jma_multilingual.xlsx （「多言語辞書（本体）」シート）

元のシートの行と列をそのまま出力しています（見出しの行は除きます）。

| 列 | 内容 | 元の項目 |
|---|---|---|
| 1 | 番号 | `No.` |
| 2 | 種別（用語・情報文・対応行動） | `種別` |
| 3 | 情報種別等 | `情報種別等` |
| 4 | 日本語 | `日本語` |
| 5 | 英語 | `英語` |
| 6 | 中国語（簡体字） | `中国語(簡体字)` |
| 7 | 中国語（繁体字） | `中国語(繁体字)` |
| 8 | 韓国語 | `韓国語` |
| 9 | ポルトガル語 | `ポルトガル語` |
| 10 | スペイン語 | `スペイン語` |
| 11 | インドネシア語 | `インドネシア語` |
| 12 | ベトナム語 | `ベトナム語` |
| 13 | タガログ語 | `タガログ語` |
| 14 | タイ語 | `タイ語` |
| 15 | ネパール語 | `ネパール語` |
| 16 | クメール語 | `クメール語` |
| 17 | ビルマ語 | `ビルマ語` |
| 18 | モンゴル語 | `モンゴル語` |
| 19 | 備考 | `備考` |
| 20 | コード表の名前（下の表を参照） | `XMLコード表` |
| 21 | コード | `コード` |
| 22 | その行の追加・修正の履歴 | `更新履歴` |

コードのある行は、次の CSV の列と結び付けられます。

| XMLコード表 | 結合先 |
|---|---|
| `AreaInformationCity` | `csv/AreaInformationCity.csv` の1列目 |
| `AreaForecastLocalM` | `csv/AreaForecastLocalM.csv` の1列目 |
| `AreaForecastLocalE` | `csv/AreaForecastLocalE_` で始まる各 CSV の1列目 |
| `AreaForecast` | `csv/AreaForecast.csv` の4列目 |
| `AreaForecastEEW` | `csv/AreaForecastEEW.csv` の1列目 |
| `AreaForecastLocalEEW` | `csv/AreaForecastLocalEEW.csv` の1列目 |
| `AreaInformationPrefectureEarthquake` | `csv/AreaInformationPrefectureEarthquake.csv` の1列目 |
| `AreaEpicenter` | `csv/AreaEpicenter.csv` の1列目 |
| `AreaEpicenterAbbreviation` | `csv/AreaEpicenterAbbreviation.csv` の1列目 |
| `AreaTsunami` | `csv/AreaTsunami.csv` の1列目 |
| `AreaFloodForecast` | `csv/AreaFloodForecast.csv` の1列目 |
| `WaterLevelStation` | `csv/WaterLevelStation.csv` の1列目 |
| `PointVolcano` | `csv/PointVolcano.csv` の1列目 |
| `PointAmedas` | `csv/AmedasRainPoint.csv` の3列目 |
| `天気予報用テロップ番号` | このリポジトリには該当する CSV がありません |

- 元データでコードが数値として入力されているため、先頭の 0 が欠けている行があります（例: 札幌中央区は `110100` で、コード表では `0110100`）。値は元データのまま出力しているので、結合する際はコード表の桁数に合わせて先頭を 0 で埋めてください。
- 同じコードに複数の行があることがあります（例: `1036600` の「上野村」と「群馬上野村」）。コードと日本語（4列目）の両方で照合してください。
- 現在のコード表に無いコード（浜松市の再編前の区など）を含みます。
- 「更新履歴」シートは出力していません。更新の内容は出典のページを参照してください。

## ライセンス

- ソースコード（`JmaCodeDictionary/`）は [MIT License](LICENSE) です。
- `csv/` 以下の CSV は、気象庁ホームページで公開されているデータを加工して作成したものです。  
  気象庁ホームページのコンテンツは [公共データ利用規約（第1.0版）](https://www.digital.go.jp/resources/open_data/public_data_license_v1.0) に準拠した利用条件の下で利用できます（[気象庁ホームページについて](https://www.jma.go.jp/jma/kishou/info/coment.html)）。  
  CSV を利用する際は、同規約に従って元データの出典と、加工されたデータであることを記載してください。

  （記載例）出典：気象庁ホームページ（https://xml.kishou.go.jp/tec_material.html）のデータを加工して作成
