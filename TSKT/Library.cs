using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace TSKT
{
    public class Library
    {
        public static Dictionary<string, Dictionary<string, string>> CreateDictionaryFromExcel(string excelPath)
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
                    foreach(var cell in row.CellsUsed().Skip(1))
                    {
                        columnLanguages.Add(
                            cell.WorksheetColumn().ColumnNumber(),
                            cell.Value.ToString());
                    }
                }

                var keyLanguageValues = new Dictionary<string, Dictionary<string, string>>();
                foreach(var row in worksheet.RowsUsed().Skip(1))
                {
                    var key = row?.Cells().First()?.Value.ToString();
                    if (key == null)
                    {
                        continue;
                    }

                    if (!keyLanguageValues.TryGetValue(key, out Dictionary<string, string> dict))
                    {
                        dict = new Dictionary<string, string>();
                        keyLanguageValues.Add(key, dict);
                    }

                    foreach (var it in columnLanguages)
                    {
                        var language = it.Value;
                        var cell = row?.Cell(it.Key);
                        var value = cell?.Value?.ToString();
                        dict.Add(language, value);
                    }
                }
                return keyLanguageValues;
            }
        }
    }
}
