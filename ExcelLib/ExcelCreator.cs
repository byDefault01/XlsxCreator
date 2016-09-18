using ExcelLib.Base;

namespace ExcelLib
{
    public static class ExcelCreator
    {
        public static void CreateExcelFile(string filePath, IFeeder feed)
        {
            var writer = new ExcelWriter(filePath);
            writer.WriteBulk(feed);
        }
    }
}
