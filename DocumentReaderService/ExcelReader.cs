using System;
using System.Collections.Generic;
using System.IO;
using DocumentReaderService.Exceptions;
using OfficeOpenXml;

namespace DocumentReaderService
{
    public class ExcelReader
    {
        private const int DocColumn = 3;
        private const int OperationColumn = 4;
        public static IDictionary<string, int> ResultDict { get; set; }

        public static IDictionary<string, int> ReadFromFile(string filePath)
        {
            var existingFile = new FileInfo(filePath);
            using (var package = new ExcelPackage(existingFile))
            {
                return GetDocumentCount(package);
            }
        }

        public static IDictionary<string, int> ReadFromByteArray(byte[] bytes)
        {
            var ms = new MemoryStream(bytes);
            using (var package = new ExcelPackage(ms))
            {
                return GetDocumentCount(package);
            }
        }

        private static IDictionary<string, int> GetDocumentCount(ExcelPackage package)
        {
            var worksheet = package.Workbook.Worksheets[1];
            var endRow = worksheet.Dimension.End.Row;
            ResultDict = new Dictionary<string, int>();
            if (!CheckFormat(worksheet))
            {
                throw new ExceptionExcelReader("Invalid file format");
            }
            for (var row = 6; row <= endRow; row++)
            {
                var documentDescription = GetCellValue(worksheet, row, DocColumn).ToString();
                var operationDescription = GetCellValue(worksheet, row, OperationColumn).ToString().ToLower();
                AddToDictionary(ResultDict, documentDescription, operationDescription);
            }

            return ResultDict;
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

        private static object GetCellValue(ExcelWorksheet worksheet, int row, int cell)
        {
            return worksheet.Cells[row, cell].Value ?? "";
        }

        private static bool CheckFormat(ExcelWorksheet worksheet)
        {
            const int headerRow = 5;
            const string headerDocument = "документ";
            const string headerOperation = "операция";
            var docColumnHeader = GetCellValue(worksheet, headerRow, DocColumn).ToString().ToLower();
            var operationColumnHeader = GetCellValue(worksheet, headerRow, OperationColumn).ToString().ToLower();
            if ( docColumnHeader == headerDocument)
            {
                return true;
            }

            if (operationColumnHeader == headerOperation)
            {
                return true;
            }

            return false;
        }
    }
}
