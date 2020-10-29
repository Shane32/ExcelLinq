namespace Shane32.ExcelLinq.Exceptions
{
    public class ColumnMissingException : InvalidDataException
    {
        public ColumnMissingException(string columnName, string sheetName) : base($"Missing column '{columnName}' in sheet '{sheetName}'") { }
    }
}
