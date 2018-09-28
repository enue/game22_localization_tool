using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

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
                var workbook = WorkbookFactory.Create(stream);
                Console.WriteLine("loaded " + excelPath);
                var worksheet = workbook.GetSheetAt(0);
                {
                    var row = worksheet.GetRow(0);
                    var cells = row.Cells;
                    for (int i = 1; i < cells.Count; ++i)
                    {
                        var value = cells[i].StringCellValue;
                        columnLanguages.Add(i, value);
                    }
                }

                var keyLanguageValues = new Dictionary<string, Dictionary<string, string>>();
                int lastRow = worksheet.LastRowNum;
                for (int i = 1; i <= lastRow; i++)
                {
                    var row = worksheet.GetRow(i);
                    var key = row?.GetCell(0)?.StringCellValue;
                    if (key == null)
                    {
                        continue;
                    }
                    foreach (var it in columnLanguages)
                    {
                        var language = it.Value;
                        var cell = row?.GetCell(it.Key);
                        var value = cell?.ToString();

                        if (!keyLanguageValues.TryGetValue(key, out Dictionary<string, string> dict))
                        {
                            dict = new Dictionary<string, string>();
                            keyLanguageValues.Add(key, dict);
                        }
                        dict.Add(language, value);
                    }
                }
                return keyLanguageValues;
            }
        }
    }
}
