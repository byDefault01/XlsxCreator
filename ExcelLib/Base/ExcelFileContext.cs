using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelLib.Base
{
    internal class ExcelFileContext : IDisposable
    {
        private SpreadsheetDocument _document;
        private WorkbookPart _workbookpart;
        private WorksheetPart _worksheetPart;
        private SheetData _sheetData;

        public SheetData SheetData { get { return _sheetData; } }

        public WorksheetPart WorksheetPart { get { return _worksheetPart; } }

        public ExcelFileContext(string filename) : this(filename, "Sheet 1")
        {
        }

        public ExcelFileContext(string filename, string sheetName)
        {
            CreateExcelFile(filename, sheetName);
        }

        private void CreateExcelFile(string filename, string sheetName)
        {
            _document = SpreadsheetDocument.Create(filename, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            _workbookpart = _document.AddWorkbookPart();
            _workbookpart.Workbook = new Workbook();
            _worksheetPart = _workbookpart.AddNewPart<WorksheetPart>();
            _sheetData = new SheetData();
            _worksheetPart.Worksheet = new Worksheet(_sheetData);

            var sheets = _document.WorkbookPart.Workbook.AppendChild(new Sheets());

            var sheet = new Sheet()
            {
                Id = _document.WorkbookPart.GetIdOfPart(_worksheetPart),
                SheetId = 1,
                Name = sheetName ?? "Sheet 1"
            };
            sheets.AppendChild(sheet);
        }

        public void Dispose()
        {
            _workbookpart.Workbook.Save();
            _document.Close();
        }
    }
}
