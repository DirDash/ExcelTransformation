using System;
using ExcelTransformation.TableClasses;
using ExcelTransformation.Utils;

namespace ExcelTransformation
{
    class Program
    {
        static void Main()
        {
            const string accountTablePostfix = "-accounts";
            const string managerTablePostfix = "-managers";
            const string relationTablePostfix = "-account-manager-relations";
            const string outputFileExtension = "xlsx";

            var initialTableUrl = GetInitialTableFileUrl();
            var accountTableUrl = GetOutputFileUrl(initialTableUrl, accountTablePostfix, outputFileExtension);
            var managerTableUrl = GetOutputFileUrl(initialTableUrl, managerTablePostfix, outputFileExtension);
            var relationTableUrl = GetOutputFileUrl(initialTableUrl, relationTablePostfix, outputFileExtension);

            var initialTable = new OpenXMLTable();
            var accountTable = new OpenXMLTable();
            var managerTable = new OpenXMLTable();
            var relationTable = new OpenXMLTable();

            var xlsNormalizer = new AccountManagerNormalizer();

            try
            {
                Console.WriteLine("Normalization in progress...");

                initialTable.Open(initialTableUrl, false);
                accountTable.Create(accountTableUrl);
                managerTable.Create(managerTableUrl);
                relationTable.Create(relationTableUrl);

                using (ExecutionTimer.StartNew("Normalization"))
                {
                    xlsNormalizer.Normalize(initialTable, accountTable, managerTable, relationTable);
                }

                initialTable.SaveAndClose();
                accountTable.SaveAndClose();
                managerTable.SaveAndClose();
                relationTable.SaveAndClose();
            }
            catch (Exception exception)
            {
                Console.WriteLine("EXCEPTION Occured:");
                Console.WriteLine(exception.Message);

                Close();
                return;
            }

            Console.WriteLine("Normalization has been successfully done.");
            Close();            
        }

        static string GetInitialTableFileUrl()
        {
            Console.WriteLine("Enter (xls|xlsx) file to normalize url: ");

            return Console.ReadLine();
        }

        static string GetOutputFileUrl(string inputTableFileName, string postfix, string fileExtension)
        {
            var outputFileUrl = string.Empty;

            var inputFileUrlSplit = inputTableFileName.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            for (var i = 0; i < inputFileUrlSplit.Length - 1; i++)
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