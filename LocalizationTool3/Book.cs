using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TSKT
{
    class Book
    {
        public class Sheet
        {
            public class Row
            {
                public List<string> Cells { get; } = new List<string>();
            }

            public string Name { get; set; } = "";
            public List<Row> Rows { get; } = new List<Row>();

            public Row AppendRow()
            {
                var row = new Row();
                Rows.Add(row);
                return row;
            }
        }
        public List<Sheet> Sheets { get; } = new List<Sheet>();

        public Book()
        {
        }

        public Book(string filename)
        {
            using var stream = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var document = SpreadsheetDocument.Open(stream, isEditable: false);
            var workbookPart = document.WorkbookPart;
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

                var s = new Sheet()
                {
                    Name = sheet.Name?.Value ?? ""
                };
                Sheets.Add(s);

                foreach (var row in worksheet.Descendants<Row>())
                {
                    var r = new Sheet.Row();
                    s.Rows.Add(r);
                    foreach (var cell in row.Descendants<Cell>())
                    {
                        TryGetCellValue(cell, sharedStringTalbePart, out var value);
                        r.Cells.Add(value);
                    }
                }
            }
        }

        public Sheet AppendSheet()
        {
            var sheet = new Sheet();
            Sheets.Add(sheet);
            return sheet;
        }

        // https://docs.microsoft.com/ja-jp/office/open-xml/how-to-retrieve-the-values-of-cells-in-a-spreadsheet
        static bool TryGetCellValue(Cell cell, SharedStringTablePart? sharedStringTablePart, out string result)
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

        public void ToXlsx(string filename)
        {
            using var document = SpreadsheetDocument.Create(filename, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            var workbookpart = document.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            SharedStringTablePart shareStringPart;
            if (workbookpart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = workbookpart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = workbookpart.AddNewPart<SharedStringTablePart>();
            }

            foreach (var it in Sheets)
            {
                var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

                uint newSheetId;
                if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Any())
                {
                    newSheetId = sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                        .Select(_ => _.SheetId.Value)
                        .Max() + 1;
                }
                else
                {
                    newSheetId = 1;
                }

                var sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = newSheetId,
                    Name = it.Name
                };
                sheets.Append(sheet);

                for (int i = 0; i < it.Rows.Count; ++i)
                {
                    var row = it.Rows[i];
                    for (int j = 0; j < row.Cells.Count; ++j)
                    {
                        var cell = row.Cells[j];
                        var c = GetCellInWorksheet((uint)j, (uint)i, worksheetPart);

                        var index = InsertSharedStringItem(cell, shareStringPart);
                        c.CellValue = new CellValue(i.ToString());
                        c.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.SharedString);
                    }
                }
            }
            document.Save();
        }

        static string CellRef(uint column, uint row)
        {
            string columnName = "";
            while (true)
            {
                var c = (char)('A' + (char)(column % 26));
                columnName = c.ToString() + columnName;

                column /= 26;
                if (column == 0)
                {
                    break;
                }
            }

            return columnName + (row + 1);
        }

        private static Cell GetCellInWorksheet(uint columnIndex, uint rowIndex, WorksheetPart worksheetPart)
        {
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            var cellReference = CellRef(columnIndex, rowIndex);

            // If the worksheet does not contain a row with the specified row index, insert one.
            var row = sheetData
                .Elements<Row>()
                .FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row()
                {
                    RowIndex = rowIndex
                };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            var cell = row.Elements<Cell>()
                .Where(c => c.CellReference.Value == cellReference)
                .FirstOrDefault();
            if (cell == null)
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                var refCell = row.Elements<Cell>()
                    .FirstOrDefault(_ => string.Compare(_.CellReference.Value, cellReference, true) > 0);

                cell = new Cell() { CellReference = cellReference };
                row.InsertBefore(cell, refCell);
            }
            return cell;
        }

        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
    }
}
