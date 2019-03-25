namespace DocumentReaderService.Settings
{
    public static class FileStructureList
    {
        public static FileStructure Evrika = new FileStructure
        {
            StartRow = 6,
            HeaderRow = 5,
            DocumentColumn = 3,
            OperationColumn = 4,
            HeaderDocumentName = "документ",
            HeaderOperationName = "операция"
        };

        public static FileStructure Fresh = new FileStructure
        {
            StartRow = 2,
            HeaderRow = 1,
            DocumentColumn = 3,
            OperationColumn = 7,
            HeaderDocumentName = "тип документа",
            HeaderOperationName = "вид операции"
        };
    }
}
