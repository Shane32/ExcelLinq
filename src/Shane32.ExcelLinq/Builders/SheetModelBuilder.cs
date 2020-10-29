using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq.Expressions;
using OfficeOpenXml;
using Shane32.ExcelLinq.Models;

namespace Shane32.ExcelLinq.Builders
{
    public class SheetModelBuilder<T> : ISheetModel, IColumnModelLookup
    {
        private readonly ExcelModelBuilder _excelModelBuilder;
        private readonly string _sheetName;
        private readonly List<string> _sheetAlternateNames = new List<string>();
        internal List<IColumnModel> _columns = new List<IColumnModel>();
        internal Dictionary<string, IColumnModel> _columnDictionary = new Dictionary<string, IColumnModel>();
        private Func<ExcelWorksheet, ExcelRange> _readRangeLocator;
        private Func<ExcelWorksheet, ExcelRange> _writeRangeLocator;
        private Action<ExcelWorksheet, ExcelRange> _writePolisher;
        private bool _optional;
        private bool _skipEmptyRows;

        public SheetModelBuilder(ExcelModelBuilder excelModelBuilder, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));
            _sheetName = name.Trim();
            _excelModelBuilder = excelModelBuilder;
            if (excelModelBuilder._typeDictionary.ContainsKey(typeof(T)))
                throw new InvalidOperationException($"Type {typeof(T).Name} already exists in the database model");
            excelModelBuilder._sheetDictionary.Add(_sheetName.ToLower(), this);
            excelModelBuilder._sheets.Add(this);
            excelModelBuilder._typeDictionary.Add(typeof(T), this);
        }

        public SheetModelBuilder<T> AlternateName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));
            name = name.Trim();
            _excelModelBuilder._sheetDictionary.Add(name.ToLower(), this);
            _sheetAlternateNames.Add(name);
            return this;
        }

        public ColumnModelBuilder<T, TReturn> Column<TReturn>(Expression<Func<T, TReturn>> memberAccessor)
        {
            if (memberAccessor.Body is MemberExpression memberExpression) {
                return Column(memberAccessor, memberExpression.Member.Name);
            }
            throw new ArgumentOutOfRangeException(nameof(memberAccessor), $"{nameof(memberAccessor)}.{nameof(LambdaExpression.Body)} must be a {nameof(MemberExpression)}");
        }

        public ColumnModelBuilder<T, TReturn> Column<TReturn>(Expression<Func<T, TReturn>> memberAccessor, string name)
        {
            if (name == null) throw new ArgumentNullException(nameof(name));
            if (memberAccessor.Body is MemberExpression memberExpression) {
                var trimmedName = name.Trim();
                for (int i = 0; i < _columns.Count; i++) {
                    if (_columns[i].Member == memberExpression.Member && _columns[i].Name == trimmedName)
                        return (ColumnModelBuilder<T, TReturn>)_columns[i];
                }
                return new ColumnModelBuilder<T, TReturn>(this, memberExpression, name);
            }
            throw new ArgumentOutOfRangeException(nameof(memberAccessor), $"{nameof(memberAccessor)}.{nameof(LambdaExpression.Body)} must be a {nameof(MemberExpression)}");
        }

        public SheetModelBuilder<T> ReadRangeLocator(Func<ExcelWorksheet, ExcelRange> readRangeLocator)
        {
            _readRangeLocator = readRangeLocator;
            return this;
        }

        public SheetModelBuilder<T> WriteRangeLocator(Func<ExcelWorksheet, ExcelRange> writeRangeLocator)
        {
            _writeRangeLocator = writeRangeLocator;
            return this;
        }

        public SheetModelBuilder<T> WritePolisher(Action<ExcelWorksheet, ExcelRange> writePolisher)
        {
            _writePolisher = writePolisher;
            return this;
        }

        public SheetModelBuilder<T> Optional()
        {
            _optional = true;
            return this;
        }

        public SheetModelBuilder<T> SkipEmptyRows()
        {
            _skipEmptyRows = true;
            return this;
        }

        bool IColumnModelLookup.TryGetValue(string columnName, out IColumnModel value)
        {
            if (string.IsNullOrWhiteSpace(columnName)) {
                value = null;
                return false;
            } else {
                return _columnDictionary.TryGetValue(columnName.Trim().ToLower(), out value);
            }
        }

        IEnumerator<IColumnModel> IEnumerable<IColumnModel>.GetEnumerator() => _columns.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => _columns.GetEnumerator();

        string ISheetModel.Name => _sheetName;

        IReadOnlyList<string> ISheetModel.AlternateNames => _sheetAlternateNames;

        IColumnModelLookup ISheetModel.Columns => this;

        Func<ExcelWorksheet, ExcelRange> ISheetModel.ReadRangeLocator => _readRangeLocator;

        Func<ExcelWorksheet, ExcelRange> ISheetModel.WriteRangeLocator => _writeRangeLocator;

        Action<ExcelWorksheet, ExcelRange> ISheetModel.WritePolisher => _writePolisher;

        Type ISheetModel.Type => typeof(T);

        bool ISheetModel.Optional => _optional;

        int IReadOnlyCollection<IColumnModel>.Count => _columns.Count;

        IColumnModel IReadOnlyList<IColumnModel>.this[int index] => _columns[index];

        IColumnModel IColumnModelLookup.this[string columnName] => string.IsNullOrWhiteSpace(columnName) ? null : _columnDictionary[columnName.Trim().ToLower()];

        bool ISheetModel.SkipEmptyRows => _skipEmptyRows;
    }
}
