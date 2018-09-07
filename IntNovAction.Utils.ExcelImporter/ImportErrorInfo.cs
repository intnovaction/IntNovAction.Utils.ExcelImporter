namespace IntNovAction.Utils.Importer
{
    public class ImportErrorInfo
    {
        public int Row { get; set; }
        public int Column { get; set; }

        public ImportErrorType ErrorType { get; set; }
    }

    public enum ImportErrorType
    {
        SheetEmpty = 1
    }
}