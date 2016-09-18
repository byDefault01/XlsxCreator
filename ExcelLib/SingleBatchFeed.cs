using System.Collections.Generic;
using System.Linq;
using ExcelLib.Base;

namespace ExcelLib
{
    public class SingleBatchFeed : IFeeder
    {
        private List<string> _headers;
        private List<List<string>> _dataRows;

        public string SheetName { get; private set; }

        public IEnumerable<string> Headers
        {
            get { return _headers; }
        }

        public IEnumerable<IEnumerable<string>> DataRows
        {
            get { return _dataRows; }
        }

        public SingleBatchFeed(string sheetName)
        {
            // todo - if the file exists, will this get automatically populated?
            _headers = new List<string>();
            _dataRows = new List<List<string>>();
            SheetName = sheetName;
        }

        public void AddHeaderColumn(string header)
        {
            _headers.Add(header);
        }

        public void AddDataRow(ICollection<string> textElements)
        {
            _dataRows.Add(textElements.ToList());
        }
    }
}
