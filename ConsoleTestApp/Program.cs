using System;
using System.IO;
using System.Linq;
using ExcelLib;

namespace ConsoleTestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"test.txt";

            var data = new SingleBatchFeed("TestSheet");
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
                        data.AddHeaderColumn(splitted[1]);

                    Console.WriteLine();
                    line = textReader.ReadLine();
                }

                ExcelCreator.CreateExcelFile(DateTime.Now.Ticks.ToString() + ".xlsx", data);
            }

            Console.ReadKey();
        }
    }
}
