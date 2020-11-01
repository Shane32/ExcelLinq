using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using OfficeOpenXml;
using Shane32.ExcelLinq.Models;

namespace Shane32.ExcelLinq.Builders
{
    public class ColumnModelBuilder<T, TReturn> : IColumnModel
    {
        private readonly SheetModelBuilder<T> _sheetModelBuilder;
        private readonly string _columnName;
        private readonly List<string> _columnAlternateNames = new List<string>();
        private readonly MemberInfo _member;
        private Func<ExcelRange, object> _readSerializer;
        private Action<ExcelRange, object> _writeSerializer;
        private Action<ExcelRange> _headerFormatter;
        private Action<ExcelRange> _columnFormatter;
        private Action<ExcelRange> _writePolisher;
        private bool _optional = false;

        public ColumnModelBuilder(SheetModelBuilder<T> sheetModelBuilder, MemberExpression memberExpression, string name)
        {
            if (sheetModelBuilder == null) throw new ArgumentNullException(nameof(sheetModelBuilder));
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));
            _member = memberExpression.Member;
            if (_member is PropertyInfo propertyInfo) {
                if (!propertyInfo.CanRead) throw new ArgumentOutOfRangeException(nameof(memberExpression), "This property cannot be read from");
                if (!propertyInfo.CanWrite) throw new ArgumentOutOfRangeException(nameof(memberExpression), "This property cannot be written from");
            // A MemberExpression can only represent a property or a field
            /*
            } else if (!(_member is FieldInfo)) {
                throw new ArgumentOutOfRangeException(nameof(memberExpression), "This member is not a property or field");
            */
            }
            if (sheetModelBuilder._columns.Any(x => x.Member == _member))
                throw new InvalidOperationException("This column has already been added to the sheet");
            _columnName = name.Trim();
            _sheetModelBuilder = sheetModelBuilder;
            sheetModelBuilder._columnDictionary.Add(_columnName.ToLower(), this);
            sheetModelBuilder._columns.Add(this);
        }

        public ColumnModelBuilder<T, TReturn> AlternateName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));
            name = name.Trim();
            _sheetModelBuilder._columnDictionary.Add(name.ToLower(), this);
            _columnAlternateNames.Add(name);
            return this;
        }

        public ColumnModelBuilder<T, TReturn> ReadSerializer(Func<ExcelRange, TReturn> readSerializer)
        {
            _readSerializer = readSerializer == null ? (Func<ExcelRange, object>)null : (excelRange) => readSerializer(excelRange);
            return this;
        }

        public ColumnModelBuilder<T, TReturn> WriteSerializer(Action<ExcelRange, TReturn> writeSerializer)
        {
            _writeSerializer = writeSerializer == null ? (Action<ExcelRange, object>)null : (excelRange, obj) => writeSerializer(excelRange, (TReturn)obj);
            return this;
        }

        public ColumnModelBuilder<T, TReturn> ColumnFormatter(Action<ExcelRange> columnFormatter)
        {
            _columnFormatter = columnFormatter;
            return this;
        }

        public ColumnModelBuilder<T, TReturn> HeaderFormatter(Action<ExcelRange> headerFormatter)
        {
            _headerFormatter = headerFormatter;
            return this;
        }

        public ColumnModelBuilder<T, TReturn> WritePolisher(Action<ExcelRange> writePolisher)
        {
            _writePolisher = writePolisher;
            return this;
        }

        public ColumnModelBuilder<T, TReturn> Optional()
        {
            _optional = true;
            return this;
        }

        string IColumnModel.Name => _columnName;

        IReadOnlyList<string> IColumnModel.AlternateNames => _columnAlternateNames;

        MemberInfo IColumnModel.Member => _member;

        Func<ExcelRange, object> IColumnModel.ReadSerializer => _readSerializer;

        Action<ExcelRange, object> IColumnModel.WriteSerializer => _writeSerializer;

        Action<ExcelRange> IColumnModel.HeaderFormatter => _headerFormatter;
        Action<ExcelRange> IColumnModel.ColumnFormatter => _columnFormatter;

        Action<ExcelRange> IColumnModel.WritePolisher => _writePolisher;

        bool IColumnModel.Optional => _optional;

        Type IColumnModel.Type => typeof(TReturn);
    }
}
