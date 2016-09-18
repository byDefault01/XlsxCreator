using System.Collections.Generic;
using ExcelLib.Base;

namespace ExcelLib
{
    class AsyncFeed : IFeeder
    {
        public delegate IEnumerable<string> GetDataMethodType();
        private GetDataMethodType _getDataRow;
        private List<string> _headers;

        public IEnumerable<IEnumerable<string>> DataRows
        {
            get
            {
                var row = _getDataRow();
                while(row != null)
                {
                    yield return row;
                    row = _getDataRow();
                }
            }
        }

        public IEnumerable<string> Headers { get { return _headers; } }

        public string SheetName { get; private set; }

        public AsyncFeed(string sheetName, GetDataMethodType getDataRowCallback)
        {
            _headers = new List<string>();
            _getDataRow = getDataRowCallback;
            SheetName = sheetName;
        }

        public void AddHeader(string header)
        {
            _headers.Add(header);
        }

        public void AddHeaderRange(IEnumerable<string> headers)
        {
            _headers.AddRange(headers);
        }
    }
}
