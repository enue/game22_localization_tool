using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Xml.Serialization;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

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

            public string key;
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

        public Sheet Distinct(bool verbose)
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
                        resultItem.pairs[index] = pair;
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

        public static Sheet? CreateFromExcel(string excelPath)
        {
            var result = new Sheet();

            using var stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var document = SpreadsheetDocument.Open(stream, isEditable: false);
            var workbookPart = document.WorkbookPart;
            if (workbookPart == null)
            {
                return null;
            }
            var sharedStringTalbePart = workbookPart.SharedStringTablePart;
            foreach (var sheet in workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>())
            {
                if (sheet == null)
                {
                    continue;
                }
                var columnLanguages = new Dictionary<int, string>();
                var worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                if (worksheetPart == null)
                {
                    continue;
                }
                var worksheet = worksheetPart.Worksheet;
                var column = 1;
                var topRow = worksheet.Descendants<Row>().First();
                foreach (Cell cell in topRow.Skip(1))
                {
                    if (TryGetCellValue(cell, sharedStringTalbePart, out var cellValue))
                    {
                        columnLanguages.Add(column, cellValue);
                    }
                    ++column;
                }

                foreach (var row in worksheet.Descendants<Row>().Skip(1))
                {
                    var cells = row.Descendants<Cell>().ToArray();

                    if (cells.Length == 0)
                    {
                        continue;
                    }
                    if (!TryGetCellValue(cells[0], sharedStringTalbePart, out var key))
                    {
                        continue;
                    }
                    if (string.IsNullOrEmpty(key))
                    {
                        continue;
                    }

                    var item = new Item();
                    result.items.Add(item);
                    item.key = key;

                    for (int i = 1; i < cells.Length; ++i)
                    {
                        if (!columnLanguages.TryGetValue(i, out var language))
                        {
                            continue;
                        }
                        var cell = cells[i];
                        if (TryGetCellValue(cell, sharedStringTalbePart, out var cellValue))
                        {
                            item.pairs.Add(new Item.Pair()
                            {
                                language = language,
                                text = cellValue,
                            });
                        }
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

        public string ToJsonString()
        {
            var json = Utf8Json.JsonSerializer.Serialize(this);
            return Utf8Json.JsonSerializer.PrettyPrint(json);
        }

        public void ToExcel()
        {
            var output = new MemoryStream();
            var document = SpreadsheetDocument.Create(output, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            var workbookpart = document.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet()
            {
                Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var row = new Row() { RowIndex = 1 };
            sheetData.Append(row);
            var cell = new Cell();
            row.Append(cell);

        }

        public XLWorkbook ToXlsx()
        {
            var book = new XLWorkbook();
            var excelSheet = book.Worksheets.Add("sheet");
            excelSheet.SheetView.Freeze(1, 1);

            var columns = new List<string>();

            {
                var index = 2;
                foreach (var item in items)
                {
                    var key = item.key;

                    var row = excelSheet.Row(index);
                    row.Cell(1).Value = key;

                    foreach (var pair in item.pairs)
                    {
                        var language = pair.language;
                        var value = pair.text;

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

            return book;
        }

        // https://docs.microsoft.com/ja-jp/office/open-xml/how-to-retrieve-the-values-of-cells-in-a-spreadsheet
        static public bool TryGetCellValue(Cell cell, SharedStringTablePart? sharedStringTablePart, out string result)
        {
            if (cell.DataType == null)
            {
                result = cell.InnerText;
                return true;
            }
            else if (cell.DataType.Value == CellValues.SharedString)
            {
                if (cell.CellValue != null && cell.CellValue.TryGetInt(out var index))
                {
                    result = sharedStringTablePart.SharedStringTable.ElementAt(index).InnerText;
                    return true;
                }
            }
            else if (cell.DataType.Value == CellValues.String)
            {
                result = cell.InnerText;
                return true;
            }
            else if (cell.DataType.Value == CellValues.Boolean)
            {
                if (cell.InnerText == "0")
                {
                    result = "FALSE";
                    return true;
                }
                else
                {
                    result = "TRUE";
                    return true;
                }
            }
            Console.WriteLine(cell.DataType.Value.ToString());

            result = "";
            return false;
        }
    }
}
