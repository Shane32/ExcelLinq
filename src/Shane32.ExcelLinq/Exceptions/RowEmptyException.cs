namespace Shane32.ExcelLinq.Exceptions
{
    public class RowEmptyException : InvalidDataException
    {
        public RowEmptyException(string sheetName) : base($"Empty row found on sheet '{sheetName}' with required columns") { }
    }
}
