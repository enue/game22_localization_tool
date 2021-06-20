using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Xml.Serialization;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

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

        public Sheet Distinct(bool verbose)
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
                        if (verbose)
                        {
                            Console.WriteLine("conflict : " + it.key + ", " + pair.language + ", [\"" + pair.text + "\" and \"" + resultItem.pairs[index].text + "\"]");
                        }
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

        public static Sheet? CreateFromExcel(string excelPath)
        {
            var result = new Sheet();

            var book = new Book(excelPath);
            foreach (var sheet in book.Sheets)
            {
                var columnLanguages = new Dictionary<int, string>();
                var column = 1;
                foreach (var cell in sheet.Rows[0].Cells.Skip(1))
                {
                    columnLanguages.Add(column, cell);
                    ++column;
                }

                foreach (var row in sheet.Rows.Skip(1))
                {
                    if (row.Cells.Count == 0)
                    {
                        continue;
                    }
                    var key = row.Cells[0];
                    if (string.IsNullOrEmpty(key))
                    {
                        continue;
                    }

                    var item = new Item();
                    result.items.Add(item);
                    item.key = key;

                    for (int i = 1; i < row.Cells.Count; ++i)
                    {
                        if (!columnLanguages.TryGetValue(i, out var language))
                        {
                            continue;
                        }
                        item.pairs.Add(new Item.Pair()
                        {
                            language = language,
                            text = row.Cells[i],
                        });
                    }
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

        public void ToXlsx(string filename)
        {
            var languageKeyTexts = CreateLanguageKeyTextDictionary();
            var languages = languageKeyTexts.Keys.ToArray();

            var book = new Book();
            var sheet = book.AppendSheet();
            sheet.Name = "sheet 1";

            var header = sheet.AppendRow();
            header.Cells.Add("");
            header.Cells.AddRange(languages);

            foreach(var it in items)
            {
                var row = sheet.AppendRow();

                row.Cells.Add(it.key);
                foreach (var lang in languages)
                {
                    var text = it.pairs.Find(_ => _.language == lang).text;
                    row.Cells.Add(text ?? "");
                }
            }

            book.ToXlsx(filename);
        }

        public XLWorkbook ToXlsxOld()
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
