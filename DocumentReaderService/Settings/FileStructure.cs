namespace DocumentReaderService.Settings
{
    public struct FileStructure
    {
        public int StartRow { get; set; }
        public int HeaderRow { get; set; }
        public int DocumentColumn { get; set; }
        public int OperationColumn { get; set; }
        public string HeaderDocumentName { get; set; }
        public string HeaderOperationName { get; set; }
    }
}
