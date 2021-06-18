using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace TSKT
{
    public class Sheet
    {
        public record Item(string Key)
        {
            readonly public List<(string language, string? text)> pairs = new();
        }

        readonly public List<Item> items = new();

        public void Merge(Sheet source)
        {
            foreach (var sourceItem in source.items)
            {
                var item = items.FirstOrDefault(_ => _.Key == sourceItem.Key);
                if (item == null)
                {
                    item = new Item(sourceItem.Key);
                    items.Add(item);
                }

                foreach (var sourcePair in sourceItem.pairs)
                {
                    var index = item.pairs.FindIndex(_ => _.language == sourcePair.language);
                    if (index < 0)
                    {
                        item.pairs.Add(sourcePair);
                    }
                    else
                    {
                        var pair = item.pairs[index];
                        Console.WriteLine("conflict : " + sourceItem.Key + ", " + sourcePair.language + ", [" + pair.text + " and " + sourcePair.text + "]");
                        item.pairs[index] = sourcePair;
                    }
                }
            }
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
                        dict = new Dictionary<string, string?>();
                        languageKeyValues.Add(language, dict);
                    }
                    dict.Add(item.Key, value);
                }
            }
            return languageKeyValues;
        }

        public static Sheet CreateFromExcel(string excelPath)
        {
            var columnLanguages = new Dictionary<int, string>();
            Console.WriteLine("load " + excelPath);
            using var stream = new System.IO.FileStream(excelPath,
                System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite);
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

                var item = new Item(key);
                result.items.Add(item);

                foreach (var it in columnLanguages)
                {
                    var language = it.Value;
                    var cell = row?.Cell(it.Key);
                    var value = cell?.Value?.ToString();

                    item.pairs.Add((language, value));
                }
            }

            return result;
        }
    }
}