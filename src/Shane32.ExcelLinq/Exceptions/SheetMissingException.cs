namespace Shane32.ExcelLinq.Exceptions
{
    public class SheetMissingException : InvalidDataException
    {
        public SheetMissingException(string sheetName) : base($"Could not find sheet '{sheetName}'") { }
    }
}
