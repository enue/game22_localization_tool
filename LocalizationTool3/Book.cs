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
                public List<string> Cells { get; } = [];
            }

            public string Name { get; set; } = "";
            public List<Row> Rows { get; } = [];

            public Row AppendRow()
            {
                var row = new Row();
                Rows.Add(row);
                return row;
            }

            public void Set(Bonn.CellReference position, string value)
            {
                while (Rows.Count < position.rowIndex)
                {
                    AppendRow();
                }
                var row = Rows[(int)position.rowIndex - 1];
                while (row.Cells.Count < position.columnIndex)
                {
                    row.Cells.Add("");
                }
                row.Cells[(int)position.columnIndex - 1] = value;
            }
        }

        public List<Sheet> Sheets { get; } = [];

        public Book()
        {
        }

        public Book(string filename)
        {
            using var stream = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var document = SpreadsheetDocument.Open(stream, isEditable: false);
            var sweetDoc = new Bonn.Document(document);
            foreach (var it in sweetDoc.Sheets)
            {
                var columnLanguages = new Dictionary<int, string>();

                var sheet = new Sheet()
                {
                    Name = it.Name ?? ""
                };
                Sheets.Add(sheet);

                foreach (var (position, value) in it.Cells())
                {
                    sheet.Set(position, value);
                }
            }
        }

        public Sheet AppendSheet()
        {
            var sheet = new Sheet();
            Sheets.Add(sheet);
            return sheet;
        }

        public void ToXlsx(string filename)
        {
            using var document = SpreadsheetDocument.Create(filename, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            var excel = new Bonn.Document(document);

            foreach (var it in Sheets)
            {
                var sheet = excel.CreateSheet(it.Name);

                for (int i = 0; i < it.Rows.Count; ++i)
                {
                    var row = it.Rows[i];
                    for (int j = 0; j < row.Cells.Count; ++j)
                    {
                        var cell = row.Cells[j];
                        var position = new Bonn.CellReference((uint)(j + 1), (uint)(i + 1));
                        sheet.SetCellValue(position, cell);
                    }
                }
            }

            document.Save();
        }

    }
}
