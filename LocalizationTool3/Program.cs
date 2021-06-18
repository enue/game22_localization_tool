﻿using System;
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
            string? outputFile = null;
            var inputFiles = new List<string>();
            for (int i = 0; i < args.Length; ++i)
            {
                if (args[i] == "in")
                {
                    inputFiles.Add(args[i + 1]);
                    ++i;
                }
                else if (args[i] == "out")
                {
                    outputFile = args[i + 1];
                    ++i;
                }
            }

            if (string.IsNullOrEmpty(outputFile))
            {
                Console.WriteLine("require out filename.");
                return;
            }

            var outputExtension = Path.GetExtension(outputFile);
            if (outputExtension == ".xlsx")
            {
                JsonsToExcel(inputFiles.ToArray(), outputFile);
            }
            else if (outputExtension == ".xml")
            {
                ExcelsToXml(inputFiles.ToArray(), outputFile);
            }
            else
            {
                ExcelsToJson(inputFiles.ToArray(), outputFile);
            }
        }

        static void ExcelsToJson(string[] excelPaths, string jsonPath)
        {
            var mergedSheet = new Sheet();
            var sheets = excelPaths.Select(_ => Sheet.CreateFromExcel(_));
            foreach (var sheet in sheets)
            {
                mergedSheet.Merge(sheet);
            }
            var json = Utf8Json.JsonSerializer.Serialize(mergedSheet);
            var prettyJson = Utf8Json.JsonSerializer.PrettyPrintByteArray(json);
            Console.WriteLine("write " + jsonPath);
            File.WriteAllBytes(jsonPath, prettyJson);
            Console.WriteLine("finished");
        }

        static void ExcelsToXml(string[] excelPaths, string xmlPath)
        {
            var mergedSheet = new Sheet();
            var sheets = excelPaths.Select(_ => Sheet.CreateFromExcel(_));
            foreach (var sheet in sheets)
            {
                mergedSheet.Merge(sheet);
            }
            var serializer = new XmlSerializer(typeof(Sheet));

            var sb = new StringBuilder();
            using (var writer = new StringWriter(sb))
            {
                serializer.Serialize(writer, mergedSheet);
            };
            var xmlString = sb.ToString();

            Console.WriteLine("write " + xmlPath);
            File.WriteAllText(xmlPath, xmlString);
            Console.WriteLine("finished");
        }

        static void JsonsToExcel(string[] jsonPaths, string excelPath)
        {
            var mergedSheet = new Sheet();
            foreach (var jsonPath in jsonPaths)
            {
                Console.WriteLine("load " + jsonPath);
                var json = File.ReadAllBytes(jsonPath);
                var sheet = Utf8Json.JsonSerializer.Deserialize<Sheet>(json);
                mergedSheet.Merge(sheet);
            }

            var book = new XLWorkbook();
            var excelSheet = book.Worksheets.Add("sheet");
            excelSheet.SheetView.Freeze(1, 1);

            var columns = new List<string>();

            {
                var index = 2;
                foreach (var item in mergedSheet.items)
                {
                    var key = item.Key;

                    var row = excelSheet.Row(index);
                    row.Cell(1).Value = key;

                    foreach (var (language, value) in item.pairs)
                    {
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


            Console.WriteLine("save " + excelPath);
            using (var fs = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                book.SaveAs(fs);
            }
            Console.WriteLine("finished");
        }
    }
}
