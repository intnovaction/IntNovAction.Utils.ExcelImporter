namespace IntNovAction.Utils.Importer
{
    public class ImportErrorInfo
    {
        public int Row { get; set; }
        public int Column { get; set; }

        public ImportErrorType ErrorType { get; set; }
        public string ColumnName { get; internal set; }
        public string CellValue { get; internal set; }
    }

    public enum ImportErrorType
    {
        SheetEmpty = 1,
        ColumnNotFound = 2,
        InvalidValue = 3
    }
}