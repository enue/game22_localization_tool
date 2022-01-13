using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using Microsoft.Extensions.CommandLineUtils;

namespace TSKT
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new CommandLineApplication(throwOnUnexpectedArg: true);
            app.HelpOption("-?|-h|--help");

            app.Command("convert", command =>
            {
                command.HelpOption("-?|-h|--help");
                var input = command.Argument("input", "input filename");
                var output = command.Argument("output", "output filename");

                command.OnExecute(() =>
                {
                    Console.WriteLine("convert " + input.Value + " to " + output.Value);
                    var sheet = ReadFile(input.Value) ?? new Sheet();
                    Write(sheet, output.Value);
                    return 0;
                });
            });
            app.Command("export", command =>
            {
                command.HelpOption("-?|-h|--help");
                var input = command.Argument("input", "filename");
                var output = command.Argument("output", "filename");
                var source = command.Argument("source", "source language(bcp47)");
                var target = command.Argument("target", "target language(bcp47)");
                var note = command.Argument("note", "note column");

                command.OnExecute(() =>
                {
                    Console.WriteLine("export xliff " + input.Value + " to " + output.Value);
                    var sheet = ReadFile(input.Value);
                    var bytes = sheet.ToXliff(source.Value, target.Value, note.Value);
                    File.WriteAllBytes(output.Value, bytes);
                    return 0;
                });
            });

            app.Command("distinct", command =>
            {
                command.HelpOption("-?|-h|--help");
                var verbose = command.Option("-v|--verbose", "alert when conflicts occur", CommandOptionType.NoValue);
                var target = command.Argument("target", "target filename");

                command.OnExecute(() =>
                {
                    Console.WriteLine("distinct " + target.Value);
                    var sheet = ReadFile(target.Value)
                        .Distinct(verbose: verbose.Value() != null);
                    Write(sheet, target.Value);
                    return 0;
                });
            });
            app.Command("rename", command =>
            {
                command.HelpOption("-?|-h|--help");
                command.Command("language", _ =>
                {
                    _.HelpOption("-?|-h|--help");
                    var target = _.Argument("file", "target filename");
                    var from = _.Argument("from", "target language");
                    var to = _.Argument("to", "new language");

                    _.OnExecute(() =>
                    {
                        Console.WriteLine("rename language " + from.Value + " to " + to.Value + " in " + target.Value);
                        var sheet = ReadFile(target.Value)
                            .RenameLanguage(from.Value, to.Value);
                        Write(sheet, target.Value);
                        return 0;
                    });
                });
            });
            app.Command("select", command =>
            {
                command.HelpOption("-?|-h|--help");
                command.Command("languages", _ =>
                {
                    _.HelpOption("-?|-h|--help");
                    var target = _.Argument("file", "target filename");
                    var languages = _.Argument("languages", "languages", multipleValues: true);

                    _.OnExecute(() =>
                    {
                        Console.WriteLine("select languages " + target.Value);
                        var sheet = ReadFile(target.Value)
                            .SelectLanguages(languages.Values.ToArray());
                        Write(sheet, target.Value);
                        return 0;
                    });
                });
            });


            app.Command("add", command =>
            {
                command.HelpOption("-?|-h|--help");
                var dest = command.Argument("dest", "destination filename", multipleValues: false);
                var sources = command.Argument("source", "source filenames", multipleValues: true);

                command.OnExecute(() =>
                {
                    Console.WriteLine("destination : " + dest.Value);
                    var sheet = ReadFile(dest.Value);
                    foreach (var it in sources.Values)
                    {
                        Console.WriteLine("add " + it);
                        var source = ReadFile(it);
                        sheet.Add(source);
                    }
                    Write(sheet, dest.Value);
                    return 0;
                });
            });
            app.Execute(args);

        }

        static Sheet ReadFile(string path)
        {
            var extension = Path.GetExtension(path);
            if (extension == ".xlsx")
            {
                return Sheet.CreateFromExcel(path);
            }
            else if (extension == ".xml")
            {
                var bytes = File.ReadAllBytes(path);
                return Sheet.CreateFromXliff(bytes);
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
                sheet.ToXlsx(path);
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
