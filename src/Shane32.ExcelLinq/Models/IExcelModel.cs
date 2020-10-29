namespace Shane32.ExcelLinq.Models
{
    public interface IExcelModel
    {
        ISheetModelLookup Sheets { get; }
        bool IgnoreSheetNames { get; }
    }
}
