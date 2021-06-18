using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace TSKT
{
    public class Sheet
    {
        public class Item
        {
            public struct Pair
            {
                public string language;
                public string? text;
            }

            public string key;
            public List<Pair> pairs = new();
        }

        public List<Item> items = new();

        public void Merge(Sheet source)
        {
            foreach (var sourceItem in source.items)
            {
                var item = new Item()
                {
                    key = sourceItem.key
                };
                items.Add(item);
                item.pairs.AddRange(sourceItem.pairs);
            }
        }

        public Sheet Compact()
        {
            var result = new Sheet();
            foreach (var it in items)
            {
                var compactPairs = it.pairs.Where(_ => !string.IsNullOrEmpty(_.text)).ToList();
                if (compactPairs.Count > 0)
                {
                    var resultItem = new Item
                    {
                        key = it.key
                    };
                    resultItem.pairs.AddRange(compactPairs);
                    result.items.Add(resultItem);
                }
            }
            return result;
        }

        public Sheet Distinct()
        {
            var result = new Sheet();
            foreach (var it in items)
            {
                var resultItem = result.items.FirstOrDefault(_ => _.key == it.key);
                if (resultItem == null)
                {
                    resultItem = new Item();
                    result.items.Add(resultItem);
                }

                foreach (var pair in it.pairs)
                {
                    var index = resultItem.pairs.FindIndex(_ => _.language == pair.language);
                    if (index < 0)
                    {
                        resultItem.pairs.Add(pair);
                    }
                    else
                    {
                        Console.WriteLine("conflict : " + it.key + ", " + pair.language + ", [" + pair.text + " and " + resultItem.pairs[index].text + "]");
                        // 後入れ優先
                        resultItem.pairs[index] = pair;
                    }
                }
            }
            return result;
        }

        public Dictionary<string, Dictionary<string, string?>> CreateLanguageKeyTextDictionary()
        {
            var languageKeyValues = new Dictionary<string, Dictionary<string, string?>>();
            foreach (var item in items)
            {
                foreach (var languageValue in item.pairs)
                {
                    var language = languageValue.language;
                    var value = languageValue.text;
                    if (!languageKeyValues.TryGetValue(language, out var dict))
                    {
                        dict = new();
                        languageKeyValues.Add(language, dict);
                    }
                    dict.Add(item.key, value);
                }
            }
            return languageKeyValues;
        }

        public static Sheet CreateFromExcel(string excelPath)
        {
            var columnLanguages = new Dictionary<int, string>();
            Console.WriteLine("load " + excelPath);
            using (var stream = new System.IO.FileStream(excelPath,
                System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
            {
                var workbook = new XLWorkbook(stream);
                Console.WriteLine("loaded " + excelPath);
                var worksheet = workbook.Worksheet(1);
                {
                    var row = worksheet.Rows().First();
                    foreach (var cell in row.CellsUsed().Skip(1))
                    {
                        columnLanguages.Add(
                            cell.WorksheetColumn().ColumnNumber(),
                            cell.Value.ToString() ?? "");
                    }
                }

                var result = new Sheet();
                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    var key = row?.Cells().First()?.Value.ToString();
                    if (key == null)
                    {
                        continue;
                    }

                    var item = new Item();
                    result.items.Add(item);
                    item.key = key;

                    foreach (var it in columnLanguages)
                    {
                        var language = it.Value;
                        var cell = row?.Cell(it.Key);
                        var value = cell?.Value?.ToString();

                        item.pairs.Add(new Item.Pair()
                        {
                            language = language,
                            text = value,
                        });
                    }
                }

                return result;
            }
        }
    }
}