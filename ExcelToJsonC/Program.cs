using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToJsonC
{
    class Program
    {
        static void Main(string[] args)
        {
            // excelを読み込んで
            // jsonにして吐き出す。

            foreach (var filename in Library.Constants.Filenames)
            {
                var columnLanguages = new Dictionary<int, string>();

                Console.WriteLine("load " + filename + ".xlsx");
                using (var stream = new System.IO.FileStream(filename + ".xlsx",
                    System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
                {
                    var workbook = WorkbookFactory.Create(stream);
                    Console.WriteLine("loaded " + filename + ".xlsx");
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
                        var key = row.GetCell(0).StringCellValue;
                        foreach (var it in columnLanguages)
                        {
                            var language = it.Value;
                            var cell = row?.GetCell(it.Key);
                            if (cell == null)
                            {
                                continue;
                            }

                            var value = cell.StringCellValue;

                            Dictionary<string, string> dict;
                            if (!keyLanguageValues.TryGetValue(key, out dict))
                            {
                                dict = new Dictionary<string, string>();
                                keyLanguageValues.Add(key, dict);
                            }
                            dict.Add(language, value);
                        }
                    }
                    var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(keyLanguageValues, Newtonsoft.Json.Formatting.Indented);
                    Console.WriteLine("write " + filename + ".json");
                    System.IO.File.WriteAllText("mst_develop_localize_Data/StreamingAssets/" + filename + ".json", jsonString);
                }
            }
            Console.WriteLine("finished");

            Console.ReadKey();
        }
    }
}
