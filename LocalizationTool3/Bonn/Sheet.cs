﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TSKT.Bonn
{
    class Sheet
    {
        readonly Document parent;
        readonly DocumentFormat.OpenXml.Spreadsheet.Sheet sheet;
        public string? Name => sheet.Name;

        Worksheet Worksheet
        {
            get
            {
                var worksheetPart = parent.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;
                if (worksheetPart == null)
                {
                    return null;
                }
                return worksheetPart.Worksheet;
            }
        }

        SheetData SheetData => Worksheet.Elements<SheetData>().First();

        public Sheet(Document parent, DocumentFormat.OpenXml.Spreadsheet.Sheet sheet)
        {
            this.parent = parent;
            this.sheet = sheet;
        }

        public Row GetOrCreateRow(uint rowIndex)
        {
            // If the worksheet does not contain a row with the specified row index, insert one.
            var row = SheetData
                .Elements<Row>()
                .FirstOrDefault(r => r.RowIndex == rowIndex);

            if (row == null)
            {
                row = new Row()
                {
                    RowIndex = rowIndex
                };
                SheetData.Append(row);
            }
            return row;
        }

        public void SetCellValue(CellReference position, string text)
        {
            var cell = GetOrCreateCell(position);
            cell.CellValue = parent.CreateCellValue(text);
            cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.SharedString);
        }

        public Cell GetOrCreateCell(CellReference cellReference)
        {
            var row = GetOrCreateRow(cellReference.rowIndex);

            // If there is not a cell with the specified column name, insert one.  
            var cell = row.Elements<Cell>()
                .Where(c => c.CellReference.Value == cellReference.value)
                .FirstOrDefault();
            if (cell == null)
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                var refCell = row.Elements<Cell>()
                    .FirstOrDefault(_ => string.Compare(_.CellReference.Value, cellReference.value, true) > 0);

                cell = new Cell() { CellReference = cellReference.value };
                row.InsertBefore(cell, refCell);
            }
            return cell;
        }

        public IEnumerable<(CellReference position, string value)> Cells()
        {
            foreach (var row in Worksheet.Descendants<Row>())
            {
                foreach (var cell in row.Descendants<Cell>())
                {
                    if (TryGetCellValue(cell, out var value))
                    {
                        var pos = new CellReference(cell.CellReference.Value);
                        yield return (pos, value);
                    }
                }
            }
        }

        // https://docs.microsoft.com/ja-jp/office/open-xml/how-to-retrieve-the-values-of-cells-in-a-spreadsheet
        bool TryGetCellValue(Cell cell, out string result)
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
                    result = parent.SharedStringTable.ElementAt(index).InnerText;
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
