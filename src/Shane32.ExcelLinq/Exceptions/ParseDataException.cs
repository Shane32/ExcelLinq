using System;

namespace Shane32.ExcelLinq.Exceptions
{
    public class ParseDataException : InvalidDataException
    {
        public ParseDataException(string cellName, string columnName, string sheetName, Exception innerException) : base($"Could not parse cell {cellName} within column '{columnName}' on sheet '{sheetName}'", innerException) { }
    }
}
