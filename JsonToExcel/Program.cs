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
            foreach (var filename in Library.Constants.JsonExcelPaths)
            {
                var excelPath = filename.Value;
                var jsonPath = filename.Key;
                Console.WriteLine("load " + jsonPath);
                var jsonString = System.IO.File.ReadAllText(jsonPath);

                var book = new NPOI.XSSF.UserModel.XSSFWorkbook();
                var sheet = book.CreateSheet("sheet");
                sheet.CreateFreezePane(0, 1, 0, 1);

                var columns = new Dictionary<string, int>();

                var json = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);
                foreach(var it in json)
                {
                    var key = it.Key;

                    var row = sheet.CreateRow(sheet.LastRowNum + 1);
                    row.CreateCell(0).SetCellValue(key);

                    foreach(var pair in (Dictionary<string, object>)it.Value)
                    {
                        var language = pair.Key;
                        var value = (string)pair.Value;
                        int column;
                        if (!columns.TryGetValue(language, out column))
                        {
                            column = columns.Count + 2;
                            columns.Add(language, column);
                        }
                        row.CreateCell(column).SetCellValue(value);
                    }
                }

                Console.WriteLine("save " + excelPath);
                using (var fs = new System.IO.FileStream(excelPath, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write))
                {
                    book.Write(fs);
                }
            }
            Console.WriteLine("finished");
        }
    }
}
