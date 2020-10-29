using Shane32.ExcelLinq.Models;
using System;
using System.Collections;
using System.Collections.Generic;

namespace Shane32.ExcelLinq.Builders
{
    public class ExcelModelBuilder : IExcelModel, ISheetModelLookup
    {
        internal List<ISheetModel> _sheets = new List<ISheetModel>();
        internal Dictionary<string, ISheetModel> _sheetDictionary = new Dictionary<string, ISheetModel>();
        internal Dictionary<Type, ISheetModel> _typeDictionary = new Dictionary<Type, ISheetModel>();
        private bool _ignoreSheetNames;

        public SheetModelBuilder<T> Sheet<T>() where T : new()
        {
            return Sheet<T>(typeof(T).Name);
        }

        public SheetModelBuilder<T> Sheet<T>(string name) where T : new()
        {
            if (name == null) throw new ArgumentNullException(nameof(name));
            var trimmedName = name.Trim();
            for (int i = 0; i < _sheets.Count; i++) {
                if (_sheets[i].Type == typeof(T) && _sheets[i].Name == trimmedName) {
                    return (SheetModelBuilder<T>)_sheets[i];
                }
            }
            return new SheetModelBuilder<T>(this, name);
        }

        public void IgnoreSheetNames()
        {
            _ignoreSheetNames = true;
        }

        IEnumerator<ISheetModel> IEnumerable<ISheetModel>.GetEnumerator() => _sheets.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => _sheets.GetEnumerator();

        ISheetModelLookup IExcelModel.Sheets => this;

        int IReadOnlyCollection<ISheetModel>.Count => _sheets.Count;

        ISheetModel IReadOnlyList<ISheetModel>.this[int index] => _sheets[index];

        ISheetModel ISheetModelLookup.this[Type type] => _typeDictionary[type];

        ISheetModel ISheetModelLookup.this[string sheetName] => string.IsNullOrWhiteSpace(sheetName) ? null : _sheetDictionary[sheetName.Trim().ToLower()];

        bool ISheetModelLookup.TryGetValue(string sheetName, out ISheetModel value)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) {
                value = null;
                return false;
            } else {
                return _sheetDictionary.TryGetValue(sheetName.Trim().ToLower(), out value);
            }
        }

        bool ISheetModelLookup.TryGetValue(Type type, out ISheetModel value) => _typeDictionary.TryGetValue(type, out value);

        internal IExcelModel Build() => this;

        bool IExcelModel.IgnoreSheetNames => _ignoreSheetNames;
    }
}
