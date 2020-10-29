using System;
using System.Collections.Generic;

namespace Shane32.ExcelLinq.Models
{
    public interface ISheetModelLookup : IReadOnlyList<ISheetModel>
    {
        ISheetModel this[string sheetName] { get; }
        ISheetModel this[Type type] { get; }
        bool TryGetValue(string sheetName, out ISheetModel value);
        bool TryGetValue(Type type, out ISheetModel value);
    }
}
