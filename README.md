# こんなときに使う

* 翻訳データをゲームに食わせるにはjson形式が都合がよい
* でもExcelファイルも使いたい
* つまり、両方の形式をいっぱつで相互変換したい

# 使い方

batファイルから起動する。  
ファイル名はプロジェクトごとにかわるので、ユーザーが適宜編集すること。  

```dos
"jsonToExcel.exe" jsonファイル名1 xlsxファイル名1 jsonファイル名2 xlsxファイル名2...
```

```dos
"jsonToExcel.exe" reverse jsonファイル名1 xlsxファイル名1 jsonファイル名2 xlsxファイル名2...
```

# jsonサンプル
`"キー" : {"言語" : "表示文字列"}`という構造。  
`comment`要素はゲームにインポートするときにでも無視すれば良い。

```json
{
  "Unit.ウォリアーE1": {
    "jpn": "ウォリアー",
    "eng": "Warrior",
    "comment": "敵として出現するウォリアークラス"
  },
  "Unit.ウォリアーE2": {
    "jpn": "ハイウォリアー",
    "comment": "敵として出現するウォリアークラス"
  },
  "Unit.ウォリアーE3": {
    "jpn": "ソルジャー",
    "comment": "敵として出現するウォリアークラス"
  },
}
```

# xlsxサンプル
　一行目が`key, 言語1, 言語2,...`という構造。

![xlsx](https://user-images.githubusercontent.com/6186357/39691102-62be6aa8-5217-11e8-9f0e-ad99071ed8f8.png)
