namespace Shane32.ExcelLinq.Exceptions
{
    public class SheetEmptyException : InvalidDataException
    {
        public SheetEmptyException(string sheetName) : base($"Sheet '{sheetName}' is empty and has required columns") { }
    }
}
