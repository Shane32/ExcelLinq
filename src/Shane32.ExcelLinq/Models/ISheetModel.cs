using System;
using System.Collections.Generic;
using OfficeOpenXml;

namespace Shane32.ExcelLinq.Models
{
    public interface ISheetModel
    {
        string Name { get; }
        Type Type { get; }
        IReadOnlyList<string> AlternateNames { get; }
        IColumnModelLookup Columns { get; }
        Func<ExcelWorksheet, ExcelRange> ReadRangeLocator { get; }
        Func<CsvRange, CsvRange> CsvReadRangeLocator { get; }
        Func<ExcelWorksheet, ExcelRange> WriteRangeLocator { get; }
        Action<ExcelWorksheet, ExcelRange> WritePolisher { get; }
        bool Optional { get; }
        bool SkipEmptyRows { get; }
    }
}
