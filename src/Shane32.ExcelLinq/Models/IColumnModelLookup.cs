using System.Collections.Generic;

namespace Shane32.ExcelLinq.Models
{
    public interface IColumnModelLookup : IReadOnlyList<IColumnModel>
    {
        IColumnModel this[string columnName] { get; }
        bool TryGetValue(string columnName, out IColumnModel value);
    }
}
