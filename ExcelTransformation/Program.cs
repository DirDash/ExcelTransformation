using System;
using ExcelTransformation.Tables;

namespace ExcelTransformation
{
    class Program
    {
        static void Main(string[] args)
        {
            const string outputFileExtension = "xls";
            
            Console.WriteLine("Enter (xls|xlsx) file url: ");
            string fileUrl = Console.ReadLine();

            var initialTable = new SpireXlsTable();
            var accountTable = new SpireXlsTable();
            var managerTable = new SpireXlsTable();
            var relationTable = new SpireXlsTable();

            var xlsNormalizer = new AccountManagerNormalizer();
            try
            {
                initialTable.LoadFromFile(fileUrl);
                xlsNormalizer.Normalize(initialTable, accountTable, managerTable, relationTable);
                accountTable.SaveToFile(GetOutputFileUrl(fileUrl, "-accounts", outputFileExtension));
                managerTable.SaveToFile(GetOutputFileUrl(fileUrl, "-managers", outputFileExtension));
                relationTable.SaveToFile(GetOutputFileUrl(fileUrl, "-account-manager-relations", outputFileExtension));
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

        static string GetOutputFileUrl(string inputTableFileName, string postfix, string fileExtension)
        {
            string outputFileUrl = string.Empty;

            string[] inputFileUrlSplit = inputTableFileName.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < inputFileUrlSplit.Length - 1; i++)
            {
                outputFileUrl += inputFileUrlSplit[i];
            }
            outputFileUrl += postfix;
            outputFileUrl += "." + fileExtension;

            return outputFileUrl;
        }
    }
}