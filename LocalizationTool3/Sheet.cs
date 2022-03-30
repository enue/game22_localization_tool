using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Xml.Serialization;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Localization.Xliff.OM.Core;

namespace TSKT
{
    public class Sheet
    {
        public class Item
        {
            public struct Pair
            {
                public string language;
                public string? text;
            }

            public string key = "";
            public List<Pair> pairs = new();
        }

        public List<Item> items = new();

        public void Add(Sheet source)
        {
            foreach (var sourceItem in source.items)
            {
                var item = new Item()
                {
                    key = sourceItem.key
                };
                items.Add(item);
                item.pairs.AddRange(sourceItem.pairs);
            }
        }

        public Sheet Distinct(bool verbose, bool selectFirstValue)
        {
            var result = new Sheet();
            foreach (var it in items)
            {
                var resultItem = result.items.FirstOrDefault(_ => _.key == it.key);
                if (resultItem == null)
                {
                    resultItem = new Item()
                    {
                        key = it.key,
                    };
                    result.items.Add(resultItem);
                }

                foreach (var pair in it.pairs)
                {
                    var index = resultItem.pairs.FindIndex(_ => _.language == pair.language);
                    if (index < 0)
                    {
                        resultItem.pairs.Add(pair);
                    }
                    else
                    {
                        if (verbose)
                        {
                            Console.WriteLine("conflict : " + it.key + ", " + pair.language + ", [\"" + pair.text + "\" and \"" + resultItem.pairs[index].text + "\"]");
                        }
                        // 後入れ優先
                        if (!selectFirstValue)
                        {
                            resultItem.pairs[index] = pair;
                        }
                    }
                }
            }
            return result;
        }

        public Sheet Trim()
        {
            var result = new Sheet();
            foreach (var it in items)
            {
                var pairs = it.pairs.Where(_ => !string.IsNullOrEmpty(_.text)).ToList();
                if (pairs.Count > 0)
                {
                    var item = new Item()
                    {
                        key = it.key,
                        pairs = pairs,
                    };
                    result.items.Add(item);
                }
            }
            return result;
        }

        public Sheet RenameLanguage(string from, string to)
        {
            var result = new Sheet();
            foreach (var it in items)
            {
                var item = new Item()
                {
                    key = it.key,
                };
                result.items.Add(item);

                foreach (var pair in it.pairs)
                {
                    if (pair.language == from)
                    {
                        item.pairs.Add(pair with { language = to });
                    }
                    else
                    {
                        item.pairs.Add(pair);
                    }
                }
            }
            return result;
        }

        public Sheet SelectLanguages(params string[] languages)
        {
            var result = new Sheet();
            foreach (var it in items)
            {
                var item = new Item()
                {
                    key = it.key,
                };
                result.items.Add(item);

                foreach (var pair in it.pairs)
                {
                    if (languages.Contains(pair.language))
                    {
                        item.pairs.Add(pair);
                    }
                }
            }
            return result;
        }

        public Dictionary<string, Dictionary<string, string?>> CreateLanguageKeyTextDictionary()
        {
            var languageKeyValues = new Dictionary<string, Dictionary<string, string?>>();
            foreach (var item in items)
            {
                foreach (var languageValue in item.pairs)
                {
                    var language = languageValue.language;
                    var value = languageValue.text;
                    if (!languageKeyValues.TryGetValue(language, out var dict))
                    {
                        dict = new();
                        languageKeyValues.Add(language, dict);
                    }
                    dict.Add(item.key, value);
                }
            }
            return languageKeyValues;
        }

        public static Sheet CreateFromExcel(string excelPath, string? sheetName)
        {
            var result = new Sheet();

            var book = new Book(excelPath);
            foreach (var sheet in book.Sheets)
            {
                if (sheetName != null)
                {
                    if (sheetName != sheet.Name)
                    {
                        continue;
                    }
                }
                var columnLanguages = new Dictionary<int, string>();
                var column = 1;
                foreach (var cell in sheet.Rows[0].Cells.Skip(1))
                {
                    columnLanguages.Add(column, cell);
                    ++column;
                }

                foreach (var row in sheet.Rows.Skip(1))
                {
                    if (row.Cells.Count == 0)
                    {
                        continue;
                    }
                    var key = row.Cells[0];
                    if (string.IsNullOrEmpty(key))
                    {
                        continue;
                    }

                    var item = new Item();
                    result.items.Add(item);
                    item.key = key;

                    foreach (var it in columnLanguages)
                    {
                        string text;
                        if (it.Key < row.Cells.Count)
                        {
                            text = row.Cells[it.Key];
                        }
                        else
                        {
                            text = "";
                        }
                        item.pairs.Add(new Item.Pair()
                        {
                            language = it.Value,
                            text = text,
                        });
                    }
                }
            }

            return result;
        }

        public string ToXmlString()
        {
            var serializer = new XmlSerializer(typeof(Sheet));

            var sb = new StringBuilder();
            using (var writer = new StringWriter(sb))
            {
                serializer.Serialize(writer, this);
            };

            return sb.ToString();
        }

        public static Sheet CreateFromJson(ReadOnlySpan<byte> json)
        {
            var option = new System.Text.Json.JsonSerializerOptions();
            option.IncludeFields = true;
            return System.Text.Json.JsonSerializer.Deserialize<Sheet>(json, option);
        }

        public string ToJsonString()
        {
            var options = new System.Text.Json.JsonSerializerOptions();
            options.WriteIndented = true;
            options.IncludeFields = true;
            options.Encoder = System.Text.Encodings.Web.JavaScriptEncoder.Create(System.Text.Unicode.UnicodeRanges.All);

            return System.Text.Json.JsonSerializer.Serialize(this, options);
        }

        public void ToXlsx(string filename)
        {
            var languageKeyTexts = CreateLanguageKeyTextDictionary();
            var languages = languageKeyTexts.Keys.ToArray();

            var book = new Book();
            var sheet = book.AppendSheet();
            sheet.Name = "sheet 1";

            var header = sheet.AppendRow();
            header.Cells.Add("");
            header.Cells.AddRange(languages);

            foreach(var it in items)
            {
                var row = sheet.AppendRow();

                row.Cells.Add(it.key);
                foreach (var lang in languages)
                {
                    var text = it.pairs.Find(_ => _.language == lang).text;
                    row.Cells.Add(text ?? "");
                }
            }

            book.ToXlsx(filename);
        }

        public byte[] ToXliff(string source, string target, string note)
        {
            var doc = new XliffDocument(source);
            doc.TargetLanguage = target;
            var file = new Localization.Xliff.OM.Core.File("file");
            doc.Files.Add(file);

            foreach (var it in items)
            {
                var unit = new Unit(it.key);
                unit.Space = Localization.Xliff.OM.Preservation.Preserve;

                var noteText = it.pairs.FirstOrDefault(_ => _.language == note).text;
                if (!string.IsNullOrEmpty(noteText))
                {
                    unit.Notes.Add(new Note(noteText));
                }

                file.Containers.Add(unit);
                var segment = new Segment();
                segment.Source = new Source(it.pairs.FirstOrDefault(_ => _.language == source).text);
                var targetText = it.pairs.FirstOrDefault(_ => _.language == target).text;
                if (!string.IsNullOrEmpty(targetText))
                {
                    segment.Target = new Target(targetText);
                }
                unit.Resources.Add(segment);
            }

            using var stream = new MemoryStream();

            var setting = new Localization.Xliff.OM.Serialization.XliffWriterSettings();
            setting.Indent = true;
            var writer = new Localization.Xliff.OM.Serialization.XliffWriter(setting);
            writer.Serialize(stream, doc);
            return stream.ToArray();
        }
        public static Sheet CreateFromXliff(byte[] bytes)
        {
            var reader = new Localization.Xliff.OM.Serialization.XliffReader();
            using var stream = new MemoryStream(bytes);
            var doc = reader.Deserialize(stream);

            var result = new Sheet();

            var queue = new Queue<TranslationContainer>();
            foreach (var it in doc.Files.SelectMany(_ => _.Containers))
            {
                queue.Enqueue(it);
            }
            while (queue.Count > 0)
            {
                var container = queue.Dequeue();
                if (container is Localization.Xliff.OM.Core.Group group)
                {
                    foreach (var it in group.Containers)
                    {
                        queue.Enqueue(it);
                    }
                }
                else if (container is Unit unit)
                {
                    var note = string.Join("\n", unit.Notes.Select(_ => _.Text));
                    foreach (var segment in unit.Resources.OfType<Segment>())
                    {
                        var item = new Item();
                        result.items.Add(item);
                        item.key = unit.Id;
                        item.pairs.Add(new Item.Pair() { language = doc.SourceLanguage, text = segment.Source.Text.OfType<ResourceStringContent>().FirstOrDefault()?.ToString() });
                        item.pairs.Add(new Item.Pair() { language = doc.TargetLanguage, text = segment.Target.Text.OfType<ResourceStringContent>().FirstOrDefault()?.ToString() });
                        if (!string.IsNullOrEmpty(note))
                        {
                            item.pairs.Add(new Item.Pair() { language = "note", text = note });
                        }
                    }
                }
            }

            return result;
        }
    }
}
