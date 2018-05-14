using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 3 && args[0] == "reverse")
            {
                JsonToExcel(args[2], args[1]);
            }
            else if (args.Length == 2)
            {
                ExcelToJson(args[0], args[1]);
            }
        }

        static void ExcelToJson(string excelPath, string jsonPath)
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
                    var key = row.GetCell(0).StringCellValue;
                    foreach (var it in columnLanguages)
                    {
                        var language = it.Value;
                        var cell = row?.GetCell(it.Key);
                        var value = cell?.ToString();

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
                Console.WriteLine("write " + jsonPath);
                System.IO.File.WriteAllText(jsonPath, jsonString);
            }
            Console.WriteLine("finished");
        }

        static void JsonToExcel(string jsonPath, string excelPath)
        {
            Console.WriteLine("load " + jsonPath);
            var jsonString = System.IO.File.ReadAllText(jsonPath);

            var book = new NPOI.XSSF.UserModel.XSSFWorkbook();
            var sheet = book.CreateSheet("sheet");
            sheet.CreateFreezePane(0, 1, 0, 1);

            var columns = new Dictionary<string, int>();

            var json = Newtonsoft.Json.JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JObject>(jsonString);
            foreach (var it in json)
            {
                var key = it.Key;

                var row = sheet.CreateRow(sheet.LastRowNum + 1);
                row.CreateCell(0).SetCellValue(key);

                foreach (var pair in (Newtonsoft.Json.Linq.JObject)it.Value)
                {
                    var language = pair.Key;
                    var value = (string)pair.Value;
                    int column;
                    if (!columns.TryGetValue(language, out column))
                    {
                        column = columns.Count + 1;
                        columns.Add(language, column);
                    }
                    row.CreateCell(column).SetCellValue(value);
                }
            }
            {
                var header = sheet.CreateRow(0);
                header.CreateCell(0).SetCellValue("key");
                foreach (var it in columns)
                {
                    header.CreateCell(it.Value).SetCellValue(it.Key);
                }
            }
            Console.WriteLine("save " + excelPath);
            using (var fs = new System.IO.FileStream(excelPath, System.IO.FileMode.Create, System.IO.FileAccess.Write))
            {
                book.Write(fs);
            }
            Console.WriteLine("finished");
        }
    }
}
