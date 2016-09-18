using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelLib.Base
{
    internal class ExcelWriter
    {
        private const int IndexOfFirstRow = 1;
        private readonly string _filePath;

        public ExcelWriter(string filePath)
        {
            _filePath = filePath;
        }

        public void WriteBulk(IFeeder data)
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

}
