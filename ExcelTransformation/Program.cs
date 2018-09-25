using System;
using ExcelTransformation.TableClasses;

namespace ExcelTransformation
{
    class Program
    {
        static void Main(string[] args)
        {
            const string accountTablePostfix = "-accounts";
            const string managerTablePostfix = "-managers";
            const string relationTablePostfix = "-account-manager-relations";
            const string outputFileExtension = "xlsx";
            
            string initialTableUrl = GetInitialTableFileUrl();
            string accountTableUrl = GetOutputFileUrl(initialTableUrl, accountTablePostfix, outputFileExtension);
            string managerTableUrl = GetOutputFileUrl(initialTableUrl, managerTablePostfix, outputFileExtension);
            string relationTableUrl = GetOutputFileUrl(initialTableUrl, relationTablePostfix, outputFileExtension);

            var xlsNormalizer = new AccountManagerNormalizer();

            try
            {
                var initialTable = new OpenXMLTable(initialTableUrl, false);
                var accountTable = new OpenXMLTable(accountTableUrl, true);
                var managerTable = new OpenXMLTable(managerTableUrl, true);
                var relationTable = new OpenXMLTable(relationTableUrl, true);

                Console.WriteLine("Normalization in progress...");
                xlsNormalizer.Normalize(initialTable, accountTable, managerTable, relationTable);

                accountTable.SaveAndClose();
                managerTable.SaveAndClose();
                relationTable.SaveAndClose();
            }
            catch(Exception exception)
            {
                Console.WriteLine(exception.Message);
                Close();
                return;
            }

            Console.WriteLine("Normalization has been successfully done.");
            Close();
        }

        static string GetInitialTableFileUrl()
        {
            Console.Write("Enter (xls|xlsx) file to normalize url: ");
            return Console.ReadLine();
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

        static void Close()
        {
            Console.WriteLine();
            Console.WriteLine("Press Enter to close the window...");
            Console.ReadLine();
        }
    }
}