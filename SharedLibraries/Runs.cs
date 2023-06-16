using ExcelExtensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedLibraries
{
    class Runs
    {
        public static void TestExcelReader_01()
        {
            var fileNameFull = @"C:\Users\HP\Downloads\tmp 1.xlsx";
            var xlReader = new XlXml.XlXmlReader(fileNameFull);
            var cellValue = xlReader.GetCellValueAsString("Sheet1", "J16");
            var colValues = xlReader.GetColumnValuesAsString("Sheet1", "N:N");
            var rowValues = xlReader.GetRowValuesAsString("Sheet1", "9:9");

            Console.WriteLine(cellValue);
            colValues.ForEach(x => Console.WriteLine(x));
            rowValues.ForEach(x => Console.WriteLine(x));
        }
    }
}
