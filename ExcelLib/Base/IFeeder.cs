using System.Collections.Generic;

namespace ExcelLib.Base
{
    public interface IFeeder
    {
        string SheetName { get; }

        IEnumerable<string> Headers { get; }

        IEnumerable<IEnumerable<string>> DataRows { get; }
    }
}
