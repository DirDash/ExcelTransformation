using System;
using Spire.Xls;

namespace ExcelTransformation
{
    public class XlsTransformator
    {
        public void Transform(string inputFileUrl)
        {
            Workbook inputWorkbook = new Workbook();
            inputWorkbook.LoadFromFile(inputFileUrl);

            Workbook outputWorkbook = new Workbook();
            var outputRows = outputWorkbook.ActiveSheet.Rows;
            
            outputRows[0].Cells[0].Value = "Id";
            outputRows[0].Cells[1].Value = "Manager";
            int outputRowIndex = 1;

            var inputRows = inputWorkbook.ActiveSheet.Rows;
            int inputRowIndex = 1;
            while (inputRowIndex < inputRows.Length && inputRows[inputRowIndex].Cells[0].Value != String.Empty)
            {
                string id = inputRows[inputRowIndex].Cells[0].Value;
                int inputCellIndex = 2;
                while (inputCellIndex < inputRows[inputRowIndex].Cells.Length && inputRows[inputRowIndex].Cells[inputCellIndex].Value != String.Empty)
                {
                    outputRows[outputRowIndex].Cells[0].Value = id;
                    outputRows[outputRowIndex].Cells[1].Value = inputRows[inputRowIndex].Cells[inputCellIndex].Value;
                    outputRowIndex++;
                    inputCellIndex++;
                }
                inputRowIndex++;
            }

            outputWorkbook.SaveToFile(GetOutputFileUrl(inputFileUrl));
        }

        private string GetOutputFileUrl(string inputFileUrl)
        {
            string outputFileUrl = String.Empty;

            string[] inputFileUrlSplit = inputFileUrl.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            int i = 0;
            for (; i < inputFileUrlSplit.Length - 1; i++)
            {
                outputFileUrl += inputFileUrlSplit[i];
            }
            outputFileUrl += "-transformed";
            outputFileUrl += ".xls";

            return outputFileUrl;
        }
    }
}