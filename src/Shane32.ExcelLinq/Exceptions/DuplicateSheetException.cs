namespace Shane32.ExcelLinq.Exceptions
{
    public class DuplicateSheetException : InvalidDataException
    {
        public DuplicateSheetException(string sheetName) : base($"Found duplicate sheet '{sheetName}'") { }
    }
}
