using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            for (int i = 0; i<args.Length / 2; ++i)
            {
                var jsonPath = args[i * 2];
                var excelPath = args[i * 2 + 1];
                Console.WriteLine("load " + jsonPath);
                var jsonString = System.IO.File.ReadAllText(jsonPath);

                var book = new NPOI.XSSF.UserModel.XSSFWorkbook();
                var sheet = book.CreateSheet("sheet");
                sheet.CreateFreezePane(0, 1, 0, 1);

                var columns = new Dictionary<string, int>();

                var json = Newtonsoft.Json.JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JObject>(jsonString);
                foreach(var it in json)
                {
                    var key = it.Key;

                    var row = sheet.CreateRow(sheet.LastRowNum + 1);
                    row.CreateCell(0).SetCellValue(key);

                    foreach(var pair in (Newtonsoft.Json.Linq.JObject)it.Value)
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
                    foreach(var it in columns)
                    {
                        header.CreateCell(it.Value).SetCellValue(it.Key);
                    }
                }
                Console.WriteLine("save " + excelPath);
                using (var fs = new System.IO.FileStream(excelPath, System.IO.FileMode.Create, System.IO.FileAccess.Write))
                {
                    book.Write(fs);
                }
            }
            Console.WriteLine("finished");
            Console.ReadKey();
        }
    }
}
