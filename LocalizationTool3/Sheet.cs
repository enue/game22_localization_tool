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
            public class Pair
            {
                public string language;
                public string text;
            }

            public string key;
            public List<Pair> pairs = new List<Pair>();
        }

        public List<Item> items = new List<Item>();

        public void Merge(Sheet source)
        {
            foreach (var sourceItem in source.items)
            {
                var item = items.FirstOrDefault(_ => _.key == sourceItem.key);
                if (item == null)
                {
                    item = new Item();
                    item.key = sourceItem.key;
                    items.Add(item);
                }

                foreach (var sourcePair in sourceItem.pairs)
                {
                    var pair = item.pairs.FirstOrDefault(_ => _.language == sourcePair.language);
                    if (pair == null)
                    {
                        pair = new Item.Pair();
                        pair.language = sourcePair.language;
                        item.pairs.Add(pair);
                    }
                    else
                    {
                        Console.WriteLine("conflict : " + sourceItem.key + ", " + sourcePair.language + ", [" + pair.text + " and " + sourcePair.text + "]");
                    }
                    pair.text = sourcePair.text;
                }
            }
        }
        public Dictionary<string, Dictionary<string, string>> CreateLanguageKeyTextDictionary()
        {
            var languageKeyValues = new Dictionary<string, Dictionary<string, string>>();
            foreach (var item in items)
            {
                foreach (var languageValue in item.pairs)
                {
                    var language = languageValue.language;
                    var value = languageValue.text;
                    if (!languageKeyValues.TryGetValue(language, out Dictionary<string, string> dict))
                    {
                        dict = new Dictionary<string, string>();
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
                            cell.Value.ToString());
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