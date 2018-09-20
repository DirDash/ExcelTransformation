using System;

namespace ExcelTransformation
{
    class Program
    {
        static void Main(string[] args)
        {
            //Example: C:\Users\DmitryB\Documents\ExcelTransformation\Examples\input2.xlsx
            Console.WriteLine("Enter (xls|xlsx) file url: ");
            string fileUrl = Console.ReadLine();

            var xlsNormalizer = new XlsNormalizer();
            try
            {
                xlsNormalizer.NormalizeFile(fileUrl);
            }
            catch(Exception exception)
            {
                Console.WriteLine(exception.Message);
                Console.WriteLine("Press Enter to close the window...");
                Console.ReadLine();
                return;
            }

            Console.WriteLine("Normalization is successfully done.");
            Console.WriteLine("Press Enter to close the window...");
            Console.ReadLine();
        }
    }
}