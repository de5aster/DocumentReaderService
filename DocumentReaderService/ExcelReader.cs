using System.Collections.Generic;
using System.IO;
using DocumentReaderService.Exceptions;
using OfficeOpenXml;

namespace DocumentReaderService
{
    public class ExcelReader
    {
        public static IDictionary<string, int> ResultDict { get; set; }

        private const int EvrikaDocColumn = 3;
        private const int EvrikaOperationColumn = 4;
        private const int EvrikaStartRow = 6;
        private const int FreshDocColumn = 3;
        private const int FreshOperationColumn = 7;
        private const int FreshStartRow = 2;

        public static IDictionary<string, int> ReadFromFile(string filePath)
        {
            var existingFile = new FileInfo(filePath);
            try
            {
                using (var package = new ExcelPackage(existingFile))
                {
                    return GetDocumentCount(package);
                }
            }
            catch (InvalidDataException)
            {
                throw new ExceptionExcelReader("File format isn't *.xlsx");
            }
        }

        public static IDictionary<string, int> ReadFromByteArray(byte[] bytes)
        {
            var ms = new MemoryStream(bytes);
            try
            {
                using (var package = new ExcelPackage(ms))
                {
                    return GetDocumentCount(package);
                }
            }
            catch (InvalidDataException)
            {
                throw new ExceptionExcelReader("File format isn't *.xlsx");
            }
        }

        private static IDictionary<string, int> GetDocumentCount(ExcelPackage package)
        {
            var worksheet = package.Workbook.Worksheets[1];           

            if (CheckEvrikaFormat(worksheet))
            {
                return GetDocuments(worksheet, EvrikaStartRow, EvrikaDocColumn, EvrikaOperationColumn);
            }

            if (CheckFreshFormat(worksheet))
            {
                return GetDocuments(worksheet, FreshStartRow, FreshDocColumn, FreshOperationColumn);
            }

            throw new ExceptionExcelReader("Invalid file format");
        }

        private static IDictionary<string, int> GetDocuments(ExcelWorksheet worksheet, int startRow, int docColumn, int operationColumn)
        {
            var resultDict = new Dictionary<string, int>();
            var endRow = worksheet.Dimension.End.Row;
            for (var row = startRow; row <= endRow; row++)
            {
                var documentDescription = GetCellValue(worksheet, row, docColumn).ToString();
                var operationDescription = GetCellValue(worksheet, row, operationColumn).ToString().ToLower();
                AddToDictionary(resultDict, documentDescription, operationDescription);
            }

            return resultDict;
        }

        private static void AddToDictionary(IDictionary<string, int> dict, string documentDescription, string operationDescription)
        {
            if (operationDescription == "не проведен") return;

            if (string.IsNullOrEmpty(documentDescription)) return;

            if (dict.ContainsKey(documentDescription))
            {
                dict[documentDescription]++;
            }
            else
            {
                dict.Add(documentDescription, 1);
            }
        }
        
        /// <summary>
        /// Check the structure of the downloaded file. Int and headers parameters are set independently.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        private static bool CheckEvrikaFormat(ExcelWorksheet worksheet)
        {
            const int headerRow = 5;
            const string headerDocument = "документ";
            const string headerOperation = "операция";
            var docColumnHeader = GetCellValue(worksheet, headerRow, EvrikaDocColumn).ToString().ToLower();
            var operationColumnHeader = GetCellValue(worksheet, headerRow, EvrikaOperationColumn).ToString().ToLower();
            if ( docColumnHeader == headerDocument && operationColumnHeader == headerOperation)
            {
                return true;
            }

            return false;
        }

        private static bool CheckFreshFormat(ExcelWorksheet worksheet)
        {
            const int headerRow = 1;
            const string headerDocument = "тип документа";
            const string headerOperation = "вид операции";
            var docColumnHeader = GetCellValue(worksheet, headerRow, FreshDocColumn).ToString().ToLower();
            var operationColumnHeader = GetCellValue(worksheet, headerRow, FreshOperationColumn).ToString().ToLower();
            if (docColumnHeader == headerDocument && operationColumnHeader == headerOperation)
            {
                return true;
            }

            return false;
        }

        private static object GetCellValue(ExcelWorksheet worksheet, int row, int cell)
        {
            return worksheet.Cells[row, cell].Value ?? "";
        }
    }
}
