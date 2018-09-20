using System;

namespace ExcelTransformation
{
    public class XlsNormalizer
    {
        private IXlsBook _inputBook;
        private IXlsBook _outputBook;

        private string _outputBookFirstColumnTitle = "Id";
        private string _outputBookSecondColumnTitle = "Manager";
        private string _outputFileSuffix = "-normalized";
        private string _outputFileExtension = "xls";

        public void NormalizeFile(IXlsBook inputBook, IXlsBook outputBook, string inputFileUrl)
        {
            _inputBook = inputBook;
            _inputBook.LoadFromFile(inputFileUrl);

            _outputBook = outputBook;

            FormatOutputBook();

            Normalize();

            _outputBook.SaveToFile(GetOutputFileUrl(inputFileUrl));
        }

        private void FormatOutputBook()
        {
            _outputBook.SetValue(0, 0, _outputBookFirstColumnTitle);
            _outputBook.SetValue(0, 1, _outputBookSecondColumnTitle);
        }

        private void Normalize()
        {
            int inputRowIndex = 1;
            int outputRowIndex = 1;
            while (!string.IsNullOrEmpty(_inputBook.GetValue(inputRowIndex, 0)))
            {
                string id = _inputBook.GetValue(inputRowIndex, 0);
                int inputCellIndex = 2;
                while (!string.IsNullOrEmpty(_inputBook.GetValue(inputRowIndex, inputCellIndex)))
                {
                    _outputBook.SetValue(outputRowIndex, 0, id);
                    _outputBook.SetValue(outputRowIndex, 1, _inputBook.GetValue(inputRowIndex, inputCellIndex));
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
    }
}