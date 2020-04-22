Excel形式の翻訳データをUnityに使いやすいjson形式に変換するツールです

# bat sample

```dos
"tool/ExcelToJson.exe" in ./source.xlsx out ./output.json
```

# source xlsx sample
|key|Japanese|English|
|:-:|:-:|:-:|
|Article_ブロンズソード|ブロンズソード|Bronze Sword|

# output json sample

```json
{
  "items": [
    {
      "key": "Article_ブロンズソード",
      "pairs": [
        {
          "language": "Japanese",
          "text": "ブロンズソード"
        },
        {
          "language": "English",
          "text": "Bronze Sword"
        }
      ]
    }
  ]
}
```

# unity class sample

```cs
[System.Serializable]
public class Sheet
{
    [System.Serializable]
    public class Item
    {
        [System.Serializable]
        public class Pair
        {
            public string language;
            public string text;
        }

        public string key;
        public List<Pair> pairs = new List<Pair>();
    }

    public List<Item> items = new List<Item>();
}
```

https://github.com/enue/Unity_TSKT_Localization/blob/master/Runtime/Sheet.cs

