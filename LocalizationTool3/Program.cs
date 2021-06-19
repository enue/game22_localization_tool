using System;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace TSKT
{
    class Program
    {
        static void Main(string[] args)
        {
            var file = args[0];

            Console.WriteLine("load " + file);
            var sheet = ReadFile(file);

            for (int i = 1; i < args.Length; ++i)
            {
                var arg = args[i];
                if (arg == "add")
                {
                    var path = args[i + 1];
                    Console.WriteLine("add " + path);
                    sheet.Add(ReadFile(path));
                    ++i;
                }
                else if (arg == "out")
                {
                    var path = args[i + 1];
                    Console.WriteLine("out " + path);
                    Write(sheet, path);
                    ++i;
                }
                else if (arg == "distinct")
                {
                    Console.WriteLine("distinct");
                    sheet = sheet.Distinct();
                }
                else
                {
                    Console.WriteLine("invalid argument : " + arg);
                }
            }
            Console.WriteLine("completed");
        }

        static Sheet ReadFile(string path)
        {
            var extension = Path.GetExtension(path);
            if (extension == ".xlsx")
            {
                return Sheet.CreateFromExcel(path);
            }
            else
            {
                var json = File.ReadAllBytes(path);
                return Utf8Json.JsonSerializer.Deserialize<Sheet>(json);
            }
        }

        static void Write(Sheet sheet, string path)
        {
            var extension = Path.GetExtension(path);
            if (extension == ".xml")
            {
                var xmlString = sheet.ToXmlString();
                File.WriteAllText(path, xmlString);
            }
            if (extension == ".xlsx")
            {
                var book = sheet.ToXlsx();
                using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
                book.SaveAs(fs);
            }
            else
            {
                var json = Utf8Json.JsonSerializer.Serialize(sheet);
                var prettyJson = Utf8Json.JsonSerializer.PrettyPrintByteArray(json);
                File.WriteAllBytes(path, prettyJson);
            }
        }
    }
}
