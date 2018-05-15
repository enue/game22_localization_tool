using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToCS
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("args -file filename -output outputFilename -ignore ignoreColumn");
                Console.ReadKey();
                return;
            }
            string excelPath = null;
            string outputPath = null;
            string ignoreColumn = null;

            for (int i = 0; i < args.Length / 2; ++i)
            {
                var key = args[i * 2];
                var value = args[i * 2 + 1];
                if (key == "-output")
                {
                    outputPath = value;
                }
                else if (key == "-file")
                {
                    excelPath = value;
                }
                else if (key == "-ignore")
                {
                    ignoreColumn = value;
                }
            }
            var codeString = GenerateCode(excelPath, ignoreColumn);

            Console.WriteLine("write " + outputPath);
            System.IO.File.WriteAllText(outputPath, codeString);
            Console.WriteLine("finished");
            Console.ReadKey();
        }

        static string GenerateCode(string excelPath, string ignoreColumn)
        {
            var keyLanguageValues = TSKT.Library.CreateDictionaryFromExcel(excelPath);

            var languageKeyValues = new Dictionary<string, Dictionary<string, string>>();
            foreach(var it in keyLanguageValues)
            {
                var key = it.Key;
                foreach(var languageValue in it.Value)
                {
                    var language = languageValue.Key;
                    var value = languageValue.Value;
                    Dictionary<string, string> dict;
                    if (!languageKeyValues.TryGetValue(language, out dict))
                    {
                        dict = new Dictionary<string, string>();
                        languageKeyValues.Add(language, dict);
                    }
                    dict.Add(key, value);
                }
            }

            var builder = new StringBuilder();
            builder.AppendLine("using System.Collections.Generic;");
            builder.AppendLine("namespace TSKT");
            builder.AppendLine("{");
            builder.AppendLine("    public static class Generated");
            builder.AppendLine("    {");
            builder.AppendLine("        public static readonly Dictionary<string, Dictionary<string, string>> languageKeyValues = new Dictionary<string, Dictionary<string, string>>()");
            builder.AppendLine("        {");

            foreach(var languageKeyValue in languageKeyValues)
            {
                var language = languageKeyValue.Key;
                if (language == ignoreColumn)
                {
                    continue;
                }
                builder.AppendLine("            {");
                builder.AppendLine("                \"" + language + "\",");
                builder.AppendLine("                new Dictionary<string, string>()");
                builder.AppendLine("                {");

                foreach (var keyValue in languageKeyValue.Value)
                {
                    if (keyValue.Value == null)
                    {
                        builder.AppendLine("                    {\"" + keyValue.Key.ToString() + "\", null},");
                    }
                    else
                    {
                        var escaledValue = keyValue.Value.Replace("\n", "\\n").Replace("\r", "\\r");
                        builder.AppendLine("                    {\"" + keyValue.Key.ToString() + "\", \"" + escaledValue + "\"},");
                    }
                }
                builder.AppendLine("                }");
                builder.AppendLine("            },");
            }

            builder.AppendLine("        };");
            builder.AppendLine("    }");
            builder.AppendLine("}");

            return builder.ToString();
        }
    }
}
