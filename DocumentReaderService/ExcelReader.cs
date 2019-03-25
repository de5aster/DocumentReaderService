using System.Collections.Generic;
using System.IO;
using DocumentReaderService.Exceptions;
using OfficeOpenXml;
using DocumentReaderService.Settings;

namespace DocumentReaderService
{
    public class ExcelReader
    {
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

            if (CheckFormat(worksheet, FileStructureList.Evrika))
            {
                return GetDocuments(worksheet, FileStructureList.Evrika);
            }

            if (CheckFormat(worksheet, FileStructureList.Fresh))
            {
                return GetDocuments(worksheet, FileStructureList.Fresh);
            }

            throw new ExceptionExcelReader("Invalid file format");
        }

        private static IDictionary<string, int> GetDocuments(ExcelWorksheet worksheet, FileStructure fs)
        {
            var resultDict = new Dictionary<string, int>();
            var endRow = worksheet.Dimension.End.Row;
            for (var row = fs.StartRow; row <= endRow; row++)
            {
                var documentDescription = GetCellValue(worksheet, row, fs.DocumentColumn).ToString();
                var operationDescription = GetCellValue(worksheet, row,fs.OperationColumn).ToString().ToLower();
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

        private static bool CheckFormat(ExcelWorksheet worksheet, FileStructure fs)
        {            
            var docColumnHeader = GetCellValue(worksheet, fs.HeaderRow, fs.DocumentColumn).ToString().ToLower();
            var operationColumnHeader = GetCellValue(worksheet, fs.HeaderRow, fs.OperationColumn).ToString().ToLower();
            if (docColumnHeader == fs.HeaderDocumentName && operationColumnHeader == fs.HeaderOperationName)
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
