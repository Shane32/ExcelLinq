namespace Shane32.ExcelLinq.Exceptions
{
    public class ColumnDataMissingException : InvalidDataException
    {
        public ColumnDataMissingException(string columnName, string sheetName) : base($"Missing required data in column '{columnName}' on sheet '{sheetName}'") { }
    }
}
