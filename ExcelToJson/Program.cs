﻿using ClosedXML.Excel;
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
            string outputFile = null;
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

            var outputExtension = System.IO.Path.GetExtension(outputFile);
            if (outputExtension == ".xlsx")
            {
                JsonsToExcel(inputFiles.ToArray(), outputFile);
            }
            else
            {
                ExcelsToJson(inputFiles.ToArray(), outputFile);
            }
        }

        static void ExcelsToJson(string[] excelPaths, string jsonPath)
        {
            var keyLanguageValues = excelPaths.Select(_ => TSKT.Library.CreateDictionaryFromExcel(_)).ToArray();
            var dict = MergeDictionary(keyLanguageValues);
            var json = Utf8Json.JsonSerializer.Serialize(dict);

            var prettyJson = Utf8Json.JsonSerializer.PrettyPrintByteArray(json);
            Console.WriteLine("write " + jsonPath);
            System.IO.File.WriteAllBytes(jsonPath, prettyJson);
            Console.WriteLine("finished");
        }

        static void JsonsToExcel(string[] jsonPaths, string excelPath)
        {
            var jsons = new List<Dictionary<string, Dictionary<string, string>>>();
            foreach (var jsonPath in jsonPaths)
            {
                var keyLanguageValues = new Dictionary<string, Dictionary<string, string>>();
                jsons.Add(keyLanguageValues);

                Console.WriteLine("load " + jsonPath);
                var jsonString = System.IO.File.ReadAllText(jsonPath);
                var json = Utf8Json.JsonSerializer.Deserialize<Dictionary<string, object>>(jsonString);
                foreach (var it in json)
                {
                    var key = it.Key;

                    foreach (var pair in (Dictionary<string, object>)it.Value)
                    {
                        var language = pair.Key;
                        var value = (string)pair.Value;

                        if (!keyLanguageValues.TryGetValue(key, out Dictionary<string, string> languageValueMap))
                        {
                            languageValueMap = new Dictionary<string, string>();
                            keyLanguageValues.Add(key, languageValueMap);
                        }
                        languageValueMap.Add(language, value);
                    }
                }
            }
            var mergedJson = MergeDictionary(jsons.ToArray());

            var book = new XLWorkbook();
            var sheet = book.Worksheets.Add("sheet");
            sheet.SheetView.Freeze(1, 1);

            var columns = new List<string>();

            {
                var index = 2;
                foreach (var keyLanguageValue in mergedJson)
                {
                    var key = keyLanguageValue.Key;

                    var row = sheet.Row(index);
                    row.Cell(1).Value = key;

                    foreach (var pair in keyLanguageValue.Value)
                    {
                        var language = pair.Key;
                        var value = pair.Value;

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
                var header = sheet.Row(1);
                header.Cell(1).Value = "key";
                var index = 2;
                foreach (var it in columns)
                {
                    header.Cell(index).Value = it;
                    ++index;
                }
            }


            Console.WriteLine("save " + excelPath);
            using (var fs = new System.IO.FileStream(excelPath, System.IO.FileMode.Create, System.IO.FileAccess.Write))
            {
                book.SaveAs(fs);
            }
            Console.WriteLine("finished");
        }

        static Dictionary<string, Dictionary<string, string>> MergeDictionary(params Dictionary<string, Dictionary<string, string>>[] dictionaries)
        {
            var merged = new Dictionary<string, Dictionary<string, string>>();
            foreach(var dictionary in dictionaries)
            {
                foreach (var it in dictionary)
                {
                    if (!merged.TryGetValue(it.Key, out Dictionary<string, string> dict))
                    {
                        dict = new Dictionary<string, string>(it.Value);
                        merged.Add(it.Key, dict);
                    }
                    else
                    {
                        foreach (var pair in it.Value)
                        {
                            dict.TryGetValue(pair.Key, out string oldValue);
                            if (string.IsNullOrEmpty(oldValue))
                            {
                                dict[pair.Key] = pair.Value;
                            }
                            else if (oldValue != pair.Value && !string.IsNullOrEmpty(pair.Value))
                            {
                                Console.WriteLine("conflict : " + it.Key + ", " + pair.Key + ", [" + oldValue + " and " + pair.Value + "]");
                            }
                        }
                    }
                }
            }
            return merged;
        }
    }
}
