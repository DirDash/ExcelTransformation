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
                Console.WriteLine("Opening initial table...");
                initialTable.Open(initialTableUrl, false);

                Console.WriteLine("Creating account table...");
                accountTable.Create(accountTableUrl);

                Console.WriteLine("Creating manager table...");
                managerTable.Create(managerTableUrl);

                Console.WriteLine("Creating relation table...");
                relationTable.Create(relationTableUrl);

                Console.WriteLine("Normalization in progress...");
                using (ExecutionTimer.StartNew("Normalization"))
                {
                    xlsNormalizer.Normalize(initialTable, accountTable, managerTable, relationTable);
                }

                Console.WriteLine("Saving...");
                initialTable.SaveAndClose();
                accountTable.SaveAndClose();
                managerTable.SaveAndClose();
                relationTable.SaveAndClose();
            }
            catch (Exception exception)
            {
                LogError(exception.Message);
                Close();
                return;
            }

            LogSuccess("Normalization has been successfully done.");
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

        static void LogError(string message)
        {
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("EXCEPTION Occured:");
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.Gray;
        }

        static void LogSuccess(string message)
        {
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.Gray;
        }

        static void Close()
        {
            Console.WriteLine();
            Console.WriteLine("Press Enter to close the window...");
            Console.ReadLine();
        }
    }
}