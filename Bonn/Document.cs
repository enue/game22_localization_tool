using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TSKT.Bonn
{
    public class Document(SpreadsheetDocument document)
    {
        readonly SpreadsheetDocument document = document;

        public WorkbookPart WorkbookPart => document.WorkbookPart ?? document.AddWorkbookPart();
        Workbook Workbook => WorkbookPart.Workbook ??= new Workbook();

        public SharedStringTable SharedStringTable
        {
            get
            {
                var part = WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()
                    ?? WorkbookPart.AddNewPart<SharedStringTablePart>();
                return part.SharedStringTable ??= new SharedStringTable();
            }
        }

        public CellValue CreateCellValue(string text)
        {
            var item = GetOrCreateSharedStringItem(text);
            return new CellValue(item);
        }

        int GetOrCreateSharedStringItem(string text)
        {
            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            int index = 0;
            foreach (var item in SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return index;
                }
                ++index;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            var result = new SharedStringItem(new Text(text));
            SharedStringTable.AppendChild(result);
            SharedStringTable.Save();

            return index;
        }

        public Sheet CreateSheet(string name)
        {
            var sheetData = new SheetData();

            var worksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(sheetData);
            var sheets = Workbook.AppendChild(new Sheets());

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
                Id = WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = newSheetId,
                Name = name
            };
            sheets.Append(sheet);

            return new Sheet(this, sheet);
        }

        public IEnumerable<Sheet> Sheets
        {
            get
            {
                foreach (var sheet in Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>())
                {
                    yield return new Sheet(this, sheet);
                }
            }
        }
    }

}
