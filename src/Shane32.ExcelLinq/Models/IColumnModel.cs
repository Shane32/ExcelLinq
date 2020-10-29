using System;
using System.Collections.Generic;
using System.Reflection;
using OfficeOpenXml;

namespace Shane32.ExcelLinq.Models
{
    public interface IColumnModel
    {
        string Name { get; }
        Type Type { get; }
        IReadOnlyList<string> AlternateNames { get; }
        MemberInfo Member { get; }
        Func<ExcelRange, object> ReadSerializer { get; }
        Action<ExcelRange, object> WriteSerializer { get; }
        Action<ExcelRange> HeaderFormatter { get; }
        Action<ExcelRange> ColumnFormatter { get; }
        Action<ExcelRange> WritePolisher { get; }
        bool Optional { get; }
    }
}
