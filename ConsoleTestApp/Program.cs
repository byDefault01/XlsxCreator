using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ConsoleTestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"test.txt";

            var data = new ExcelData("TestSheet");
            using (var textReader = new StreamReader(path))
            {
                var line = textReader.ReadLine();
                while (line != null)
                {
                    var splitted = line.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var s in splitted)
                    {
                        Console.WriteLine(s);
                    }

                    // a doua coloana vrem sa devina linie
                    if (splitted.Count() > 1)
                        data.AddHeader(splitted[1]);

                    Console.WriteLine();
                    line = textReader.ReadLine();
                }

                ExcelWriter.CreateExcelFile(DateTime.Now.Ticks.ToString() + ".xlsx", data);
            }

            Console.ReadKey();
        }
    }

    public class ExcelWriter
    {
        private const int IndexOfFirstRow = 1;
        private readonly string _filePath;

        public static void CreateExcelFile(string fileName, ExcelData data)
        {
            var writer = new ExcelWriter(fileName);
            writer.WriteBulk(data);
        }

        public ExcelWriter(string filePath)
        {
            _filePath = filePath;
        }

        public void WriteBulk(ExcelData data)
        {
            // todo - if i create an instance of this, and call this method several times, there will be some unexpected behaviour
            using (var file = new ExcelFileContext(_filePath, data.SheetName))
            {
                var sheetData = file.SheetData;

                uint rowIndex = IndexOfFirstRow;
                if (data.Headers.Any())
                {
                    // add header
                    AddRow(data.Headers, ref sheetData, rowIndex++);

                    // add column config
                    if (data.ColumnConfigurations != null)
                        AddColumnConfig(data, file.WorksheetPart);
                }

                // add each data row
                foreach (var rowData in data.DataRows)
                    AddRow(rowData, ref sheetData, rowIndex++);
            }
        }

        private static void AddRow(IEnumerable<string> textElements, ref SheetData sheetData, uint rowIndex)
        {
            var row = new Row { RowIndex = rowIndex };
            sheetData.AppendChild(row);

            AddDataToRow(ref row, textElements);
        }

        private static void AddDataToRow(ref Row row, IEnumerable<string> textElements)
        {
            var cellIndex = 0;
            foreach (var elem in textElements)
            {
                var columnLetter = GetColumnLetter(cellIndex++);
                var textCell = CreateTextCell(columnLetter, row.RowIndex, elem ?? string.Empty);
                row.AppendChild(textCell);
            }
        }

        private static void AddColumnConfig(ExcelData data, WorksheetPart worksheetPart)
        {
            var columns = (Columns)data.ColumnConfigurations.Clone();
            var sheetProperties = worksheetPart.Worksheet.SheetFormatProperties;
            worksheetPart.Worksheet.InsertAfter(columns, sheetProperties);
        }

        private static string GetColumnLetter(int intCol)
        {
            var intFirstLetter = ((intCol) / 676) + 64;
            var intSecondLetter = ((intCol % 676) / 26) + 64;
            var intThirdLetter = (intCol % 26) + 65;

            var firstLetter = (intFirstLetter > 64)
                ? (char)intFirstLetter : ' ';
            var secondLetter = (intSecondLetter > 64)
                ? (char)intSecondLetter : ' ';
            var thirdLetter = (char)intThirdLetter;

            return string.Concat(firstLetter, secondLetter,
                thirdLetter).Trim();
        }

        private static Cell CreateTextCell(string columnLetter, uint rowIndex, string cellText)
        {
            // create text object
            var text = new Text
            {
                Text = cellText
            };

            // create the inline string which will contain the text object
            var inlineString = new InlineString();
            inlineString.AppendChild(text);

            // create the cell which will contain the text object
            var cell = new Cell
            {
                DataType = CellValues.InlineString,
                CellReference = columnLetter + rowIndex
            };
            cell.AppendChild(inlineString);

            return cell;
        }
    }

    public class ExcelFileContext : IDisposable
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

    public class ExcelData
    {
        private List<string> _headers;
        private List<List<string>> _dataRows;

        // todo - I don't know exactly what this is
        public Columns ColumnConfigurations { get; set; }

        public string SheetName { get; private set; }

        public IEnumerable<string> Headers
        {
            get { return _headers; }
        }

        public IEnumerable<IEnumerable<string>> DataRows
        {
            get { return _dataRows; }
        }

        public ExcelData(string sheetName)
        {
            // todo - if the file exists, will this get automatically populated?
            _headers = new List<string>();
            _dataRows = new List<List<string>>();
            SheetName = sheetName;
        }

        public void AddHeader(string header)
        {
            _headers.Add(header);
        }

        public void AddDataRow(ICollection<string> textElements)
        {
            _dataRows.Add(textElements.ToList());
        }
    }
}
