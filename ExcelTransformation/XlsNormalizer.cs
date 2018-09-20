using System;
using Spire.Xls;

namespace ExcelTransformation
{
    public class XlsNormalizer
    {
        private string _outputBookFirstColumnTitle = "Id";
        private string _outputBookSecondColumnTitle = "Manager";
        private string _outputFileSuffix = "-normalized";
        private string _outputFileExtension = "xls";

        public void NormalizeFile(string inputFileUrl)
        {
            var inputBook = LoadBookFromUrl(inputFileUrl);
            var outputBook = CreateOutputBook();

            NormalizeSheet(inputBook.Worksheets[0], outputBook.Worksheets[0]);

            SaveOutputBook(outputBook, GetOutputFileUrl(inputBook.FileName));
        }

        private Workbook LoadBookFromUrl(string fileUrl)
        {
            var book = new Workbook();
            book.LoadFromFile(fileUrl);
            return book;
        }

        private Workbook CreateOutputBook()
        {
            var outputBook = new Workbook();
            var firstSheet = outputBook.Worksheets[0];

            firstSheet.Rows[0].Cells[0].Value = _outputBookFirstColumnTitle;
            firstSheet.Rows[0].Cells[1].Value = _outputBookSecondColumnTitle;

            return outputBook;
        }

        private void NormalizeSheet(Worksheet inputSheet, Worksheet outputSheet)
        {
            int inputRowIndex = 1;
            int outputRowIndex = 1;
            while (inputRowIndex < inputSheet.Rows.Length && inputSheet.Rows[inputRowIndex].Cells[0].HasNumber)
            {
                string id = inputSheet.Rows[inputRowIndex].Cells[0].Value;
                int inputCellIndex = 2;
                while (inputCellIndex < inputSheet.Rows[inputRowIndex].Cells.Length && inputSheet.Rows[inputRowIndex].Cells[inputCellIndex].HasString)
                {
                    outputSheet.Rows[outputRowIndex].Cells[0].Value = id;
                    outputSheet.Rows[outputRowIndex].Cells[1].Value = inputSheet.Rows[inputRowIndex].Cells[inputCellIndex].Value;
                    outputRowIndex++;
                    inputCellIndex++;
                }
                inputRowIndex++;
            }
        }

        private string GetOutputFileUrl(string inputBookFileName)
        {
            string outputFileUrl = string.Empty;

            string[] inputFileUrlSplit = inputBookFileName.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < inputFileUrlSplit.Length - 1; i++)
            {
                outputFileUrl += inputFileUrlSplit[i];
            }
            outputFileUrl += _outputFileSuffix;
            outputFileUrl += "." + _outputFileExtension;

            return outputFileUrl;
        }

        private void SaveOutputBook(Workbook outputBook, string fileName)
        {
            outputBook.SaveToFile(fileName);
        }
    }
}