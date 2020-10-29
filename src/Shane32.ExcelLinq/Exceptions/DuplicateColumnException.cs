namespace Shane32.ExcelLinq.Exceptions
{
    public class DuplicateColumnException : InvalidDataException
    {
        public DuplicateColumnException(string columnName, string sheetName) : base($"Duplicate column '{columnName}' detected in sheet '{sheetName}") { }
    }
}
