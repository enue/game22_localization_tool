using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Xml.Serialization;
using System.IO;

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

        public void Add(Sheet source)
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

        public Sheet Distinct()
        {
            var result = new Sheet();
            foreach (var it in items)
            {
                var resultItem = result.items.FirstOrDefault(_ => _.key == it.key);
                if (resultItem == null)
                {
                    resultItem = new Item()
                    {
                        key = it.key,
                    };
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
                        Console.WriteLine("conflict : " + it.key + ", " + pair.language + ", [\"" + pair.text + "\" and \"" + resultItem.pairs[index].text + "\"]");
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
            using var stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var workbook = new XLWorkbook(stream);
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

        public string ToXmlString()
        {
            var serializer = new XmlSerializer(typeof(Sheet));

            var sb = new StringBuilder();
            using (var writer = new StringWriter(sb))
            {
                serializer.Serialize(writer, this);
            };

            return sb.ToString();
        }

        public string ToJsonString()
        {
            var json = Utf8Json.JsonSerializer.Serialize(this);
            return Utf8Json.JsonSerializer.PrettyPrint(json);
        }

        public XLWorkbook ToXlsx()
        {
            var book = new XLWorkbook();
            var excelSheet = book.Worksheets.Add("sheet");
            excelSheet.SheetView.Freeze(1, 1);

            var columns = new List<string>();

            {
                var index = 2;
                foreach (var item in items)
                {
                    var key = item.key;

                    var row = excelSheet.Row(index);
                    row.Cell(1).Value = key;

                    foreach (var pair in item.pairs)
                    {
                        var language = pair.language;
                        var value = pair.text;

                        var column = columns.IndexOf(language);
                        if (column < 0)
                        {
                            column = columns.Count;
                            columns.Add(language);
                        }

                        row.Cell(column + 2).Value = value;
                    }
                    ++index;
                }
            }
            {
                var header = excelSheet.Row(1);
                header.Cell(1).Value = "key";
                var index = 2;
                foreach (var it in columns)
                {
                    header.Cell(index).Value = it;
                    ++index;
                }
            }

            return book;
        }
    }
}
