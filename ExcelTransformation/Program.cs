using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;

namespace ExcelTransformation
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "toTransform.xlsx";
            string outputPath = "result.xlsx";

            Workbook inputWorkbook = new Workbook();
            inputWorkbook.LoadFromFile(inputPath);

            Workbook outputWorkbook = new Workbook();
            outputWorkbook.LoadFromFile(outputPath);

            Worksheet sheet = outputWorkbook.Worksheets[0];

            sheet.Range["A1"].Text = "Hello,World!";

            outputWorkbook.Save();
        }
    }
}