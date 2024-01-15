using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using OfficeOpenXml;
using Shane32.ExcelLinq.Builders;
using Shane32.ExcelLinq.Exceptions;
using Shane32.ExcelLinq.Models;
using NotVisualBasic.FileIO;
using System.Globalization;

namespace Shane32.ExcelLinq
{
    public abstract class ExcelContext
    {
        private readonly IExcelModel _model;
        private readonly List<IList> _sheets;
        private readonly Dictionary<string, int> _sheetNameLookup;
        private readonly Dictionary<Type, int> _typeLookup;
        private readonly bool _initialized;

        protected ExcelContext()
        {
            var excelModel = new ExcelModelBuilder();
            OnModelCreating(excelModel);
            _model = excelModel.Build();
            _sheets = new List<IList>(Model.Sheets.Count);
            _sheetNameLookup = new Dictionary<string, int>(Model.Sheets.Count);
            _typeLookup = new Dictionary<Type, int>(Model.Sheets.Count);
            for (int i = 0; i < Model.Sheets.Count; i++) {
                var sheet = Model.Sheets[i];
                _sheets.Add(CreateListForSheet(sheet.Type));
                _sheetNameLookup.Add(sheet.Name, i);
                foreach (var sheetName in sheet.AlternateNames)
                    _sheetNameLookup.Add(sheetName, i);
                if (!_typeLookup.ContainsKey(sheet.Type))
                    _typeLookup.Add(sheet.Type, i);
            }
            _initialized = true;
        }

        // used by unit tests only
        internal ExcelContext(IExcelModel model)
        {
            _model = model ?? throw new ArgumentNullException(nameof(model));
            _sheets = new List<IList>(Model.Sheets.Count);
            _sheetNameLookup = new Dictionary<string, int>(Model.Sheets.Count);
            _typeLookup = new Dictionary<Type, int>(Model.Sheets.Count);
            for (int i = 0; i < Model.Sheets.Count; i++) {
                var sheet = Model.Sheets[i];
                _sheets.Add(CreateListForSheet(sheet.Type));
                _sheetNameLookup.Add(sheet.Name, i);
                foreach (var sheetName in sheet.AlternateNames)
                    _sheetNameLookup.Add(sheetName, i);
                if (!_typeLookup.ContainsKey(sheet.Type))
                    _typeLookup.Add(sheet.Type, i);
            }
            _initialized = true;
        }

        public IExcelModel Model => _model ?? throw new InvalidOperationException("This instance has not yet been initialized");

        protected ExcelContext(string filename) : this()
        {
            using var stream = new FileStream(filename ?? throw new ArgumentNullException(nameof(filename)), FileMode.Open, FileAccess.Read, FileShare.Read);
            using var package = new ExcelPackage(stream);
            package.Compatibility.IsWorksheets1Based = false;
            _initialized = false;
            _sheets = InitializeReadFile(package);
            _initialized = true;
        }

        //internal ExcelContext(IExcelModel model, string filename) : this(model)
        //{
        //    if (filename == null) throw new ArgumentNullException(nameof(filename));
        //    using var stream = new FileStream(filename ?? throw new ArgumentNullException(nameof(filename)), FileMode.Open, FileAccess.Read, FileShare.Read);
        //    using var package = new ExcelPackage(stream);
        //    _initialized = false;
        //    _sheets = InitializeReadFile(package);
        //    _initialized = true;
        //}

        protected ExcelContext(Stream stream) : this()
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));
            using var package = new ExcelPackage(stream);
            package.Compatibility.IsWorksheets1Based = false;
            _initialized = false;
            _sheets = InitializeReadFile(package);
            _initialized = true;
        }

        //internal ExcelContext(IExcelModel model, Stream stream) : this(model)
        //{
        //    if (stream == null) throw new ArgumentNullException(nameof(stream));
        //    using var package = new ExcelPackage(stream);
        //    _initialized = false;
        //    _sheets = InitializeReadFile(package);
        //    _initialized = true;
        //}

        protected ExcelContext(ExcelPackage excelPackage) : this()
        {
            _initialized = false;
            _sheets = InitializeReadFile(excelPackage);
            _initialized = true;
        }

        public void ReadCsv<T>(Stream stream, IFormatProvider formatProvider = null)
        {
            if (!_initialized)
                throw new InvalidOperationException();
            var index = _typeLookup[typeof(T)];
            _sheets[index] = (List<T>)OnReadCsv(stream, Model.Sheets[index], formatProvider ?? CultureInfo.CurrentCulture);
        }

        public void ReadCsv<T>(Stream stream, string name, IFormatProvider formatProvider = null)
        {
            if (!_initialized)
                throw new InvalidOperationException();
            var index = _sheetNameLookup[name];
            _sheets[index] = (List<T>)OnReadCsv(stream, Model.Sheets[index], formatProvider ?? CultureInfo.CurrentCulture);
        }

        public static TContext OpenCsv<TContext>(Stream stream, IFormatProvider formatProvider = null)
            where TContext : ExcelContext, new()
        {
            var context = new TContext();
            if (context._sheets.Count != 1)
                throw new InvalidOperationException("Cannot open CSV file with multiple sheets.");
            context.OnReadCsv(stream, context.Model.Sheets[0], formatProvider ?? CultureInfo.CurrentCulture);
            return context;
        }

        private List<IList> InitializeReadFile(ExcelPackage excelFile)
        {
            if (excelFile == null)
                throw new ArgumentNullException(nameof(excelFile));
            var data = OnReadFile(excelFile.Workbook)
                ?? throw new InvalidOperationException("No data returned from OnReadFile");
            if (data.Count != _sheets.Count)
                throw new InvalidOperationException("Invalid number of sheets returned from OnReadFile");
            for (int i = 0; i < _sheets.Count; i++) {
                if (data[i] == null)
                    throw new InvalidOperationException($"No data returned for sheet {i}");
                if (data[i].GetType() != _sheets[i].GetType())
                    throw new InvalidOperationException($"Received sheet data type {data[i].GetType()} for sheet {i}; expected {_sheets[i].GetType()}");
            }
            return data;
        }

        protected IList CreateListForSheet(Type type)
        {
            var constructedType = typeof(List<>).MakeGenericType(new[] { type });
            return (IList)Activator.CreateInstance(constructedType);
        }

        protected IList CreateListForSheet(Type type, int capacity)
        {
            var constructedType = typeof(List<>).MakeGenericType(new[] { type });
            var constructor = constructedType.GetConstructor(new Type[] { typeof(int) });
            return (IList)constructor.Invoke(new object[] { capacity });
        }

        protected abstract void OnModelCreating(ExcelModelBuilder modelBuilder);



        /// <summary>
        /// Parses an <see cref="ExcelPackage"/> and returns all of the data within all the worksheets.
        /// <br/><br/>
        /// Optional worksheets must be included in the result as an empty list of rows.
        /// </summary>
        protected virtual List<IList> OnReadFile(ExcelWorkbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            var sheets = new List<IList>(new IList[Model.Sheets.Count]);

            var sheetArray = Model.Sheets.ToList();
            if (Model.IgnoreSheetNames) {
                int i = 0;
                foreach (var worksheet in workbook.Worksheets) {
                    var sheetModel = Model.Sheets[i];
                    var sheetData = OnReadSheet(worksheet, sheetModel);
                    sheets[i++] = sheetData ?? throw new InvalidOperationException($"{nameof(OnReadSheet)} returned null for sheet '{sheetModel.Name}'");
                }
            } else {
                foreach (var workSheet in workbook.Worksheets) {
                    if (Model.Sheets.TryGetValue(workSheet.Name, out var sheetModel)) {
                        var sheetIndex = sheetArray.IndexOf(sheetModel);
                        if (sheets[sheetIndex] != null)
                            throw new DuplicateSheetException(sheetModel.Name);
                        var sheetData = OnReadSheet(workSheet, sheetModel);
                        sheets[sheetIndex] = sheetData ?? throw new InvalidOperationException($"{nameof(OnReadSheet)} returned null for sheet '{sheetModel.Name}'");
                    }
                }
            }

            for (int i = 0; i < Model.Sheets.Count; i++) {
                if (sheets[i] == null) {
                    var sheetModel = Model.Sheets[i];
                    if (sheetModel.Optional)
                        sheets[i] = CreateListForSheet(sheetModel.Type);
                    else
                        throw new SheetMissingException(sheetModel.Name);
                }
            }

            return sheets;
        }

        /// <summary>
        /// Parses a CSV file and returns all of the data within the sheet.
        /// </summary>
        protected virtual IList OnReadCsv(Stream stream, ISheetModel model, IFormatProvider formatProvider)
        {
            using var parser = new CsvTextFieldParser(stream);

            parser.Delimiters = new string[] { "," };
            parser.HasFieldsEnclosedInQuotes = true;

            var headerRow = 0;
            var currentRow = 0;
            string[] headers = null;
            while (!parser.EndOfData) {
                if (currentRow == headerRow) {
                    headers = parser.ReadFields();
                    ++currentRow;
                    break;
                }
                ++currentRow;
            }

            IList data = CreateListForSheet(model.Type, 0);
            var columnMapping = new IColumnModel[model.Columns.Count];
            var columnMapped = new bool[model.Columns.Count];
            if (headers != null) {
                var modelColumns = model.Columns.ToList();

                for (int colIndex = 0; colIndex < headers.Length; colIndex++) {
                    var cell = headers[colIndex];
                    if (cell != null) {
                        var headerName = cell;
                        if (model.Columns.TryGetValue(headerName, out var columnModel)) {
                            var columnModelIndex = modelColumns.IndexOf(columnModel);
                            if (columnMapped[columnModelIndex])
                                throw new DuplicateColumnException(columnModel.Name, model.Name);
                            columnMapped[columnModelIndex] = true;
                            columnMapping[colIndex] = columnModel;
                        }
                    }
                }
            }

            for (int i = 0; i < model.Columns.Count; i++) {
                if (columnMapped[i] == false && !model.Columns[i].Optional)
                    throw new ColumnMissingException(model.Columns[i].Name, model.Name);
            }

            while (!parser.EndOfData) {
                var range = parser.ReadFields();
                if (range.Length != columnMapping.Length) {
                    Array.Resize(ref range, columnMapping.Length);
                }
                var obj = OnReadCSVRow(range, model, columnMapping, formatProvider);
                if (obj != null)
                    data.Add(obj);
                ++currentRow;
            }


            return data;
        }

        /// <summary>
        /// Reads a worksheet and returns a set of <see cref="List{T}"/> of the entries.
        /// <br/><br/>
        /// Must not return null.
        /// </summary>
        protected virtual IList OnReadSheet(ExcelWorksheet worksheet, ISheetModel model)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            if (model == null)
                throw new ArgumentNullException(nameof(model));
            ExcelRange dataRange = (model.ReadRangeLocator ?? DefaultReadRangeLocator)(worksheet);
            if (dataRange == null) {
                //no data on sheet
                if (model.Columns.Any(x => !x.Optional))
                    throw new SheetEmptyException(model.Name);
                return CreateListForSheet(model.Type);
            }
            IList data = CreateListForSheet(model.Type, dataRange.Rows - 1);
            var headerRow = dataRange.Start.Row;
            var firstCol = dataRange.Start.Column;
            var columns = dataRange.Columns;
            var firstRow = dataRange.Start.Row + 1;
            var lastRow = dataRange.End.Row;
            var columnMapping = new IColumnModel[columns];
            var columnMapped = new bool[model.Columns.Count];
            var modelColumns = model.Columns.ToList();

            for (int colIndex = 0; colIndex < columns; colIndex++) {
                var col = colIndex + firstCol;
                var cell = worksheet.Cells[headerRow, col];
                if (cell.Value != null) {
                    var headerName = cell.Text;
                    if (model.Columns.TryGetValue(headerName, out var columnModel)) {
                        var columnModelIndex = modelColumns.IndexOf(columnModel);
                        if (columnMapped[columnModelIndex])
                            throw new DuplicateColumnException(columnModel.Name, model.Name);
                        columnMapped[columnModelIndex] = true;
                        columnMapping[colIndex] = columnModel;
                    }
                }
            }

            for (int i = 0; i < model.Columns.Count; i++) {
                if (columnMapped[i] == false && !model.Columns[i].Optional)
                    throw new ColumnMissingException(model.Columns[i].Name, model.Name);
            }

            for (int row = firstRow; row <= lastRow; row++) {
                var range = worksheet.Cells[row, firstCol, row, firstCol + columns - 1];
                var obj = OnReadRow(range, model, columnMapping);
                if (obj != null)
                    data.Add(obj);
            }

            return data;
        }

        protected virtual object OnReadCSVRow(string[] range, ISheetModel model, IColumnModel[] columnMapping, IFormatProvider formatProvider)
        {
            if (range == null)
                throw new ArgumentNullException(nameof(range));
            if (model == null)
                throw new ArgumentNullException(nameof(range));
            if (columnMapping == null)
                throw new ArgumentNullException(nameof(columnMapping));
            var firstCol = 0;
            var columns = range.Length;
            //if (range.Rows != 1)
            //    throw new ArgumentOutOfRangeException(nameof(range), "Range must represent a single row of data");
            if (columns != columnMapping.Length)
                throw new ArgumentOutOfRangeException(nameof(columnMapping), "Number of columns in range does not match size of columnMapping array");
            var obj = Activator.CreateInstance(model.Type);
            if (range.Any(x => x != null)) {
                for (int colIndex = 0; colIndex < columns; colIndex++) {
                    var col = colIndex + firstCol;
                    var columnModel = columnMapping[colIndex];
                    if (columnModel != null) {
                        var cell = range[col]; // note that range[] resets range.Address to equal the new address
                        if (string.IsNullOrEmpty(cell)) {
                            if (!columnModel.Optional)
                                throw new ColumnDataMissingException(columnModel.Name, model.Name);
                        } else {
                            object value;
                            try {
                                //if (columnModel.ReadSerializer != null) {
                                //    value = columnModel.ReadSerializer(cell);
                                //} else {
                                value = DefaultCsvReadSerializer(cell, cell, columnModel.Type, formatProvider);
                                //}
                            } catch (Exception e) {
                                throw new ParseDataException(cell, columnModel.Name, model.Name, e);
                            }
                            if (value != null) {
                                if (columnModel.Member is PropertyInfo propertyInfo) {
                                    propertyInfo.SetMethod.Invoke(obj, new[] { value });
                                } else if (columnModel.Member is FieldInfo fieldInfo) {
                                    fieldInfo.SetValue(obj, value);
                                }
                            }
                        }
                    }
                }
            } else {
                if (model.SkipEmptyRows) {
                    obj = null;
                } else {
                    foreach (var columnModel in columnMapping) {
                        if (!columnModel.Optional)
                            throw new RowEmptyException(model.Name);
                    }
                }
            }
            return obj;
        }
        /// <summary>
        /// Parses a row of data or returns null if the row should be skipped
        /// </summary>
        protected virtual object OnReadRow(ExcelRange range, ISheetModel model, IColumnModel[] columnMapping)
        {
            if (range == null)
                throw new ArgumentNullException(nameof(range));
            if (model == null)
                throw new ArgumentNullException(nameof(range));
            if (columnMapping == null)
                throw new ArgumentNullException(nameof(columnMapping));
            var firstCol = range.Start.Column;
            var row = range.Start.Row;
            var columns = range.Columns;
            if (range.Rows != 1)
                throw new ArgumentOutOfRangeException(nameof(range), "Range must represent a single row of data");
            if (columns != columnMapping.Length)
                throw new ArgumentOutOfRangeException(nameof(columnMapping), "Number of columns in range does not match size of columnMapping array");
            var obj = Activator.CreateInstance(model.Type);
            if (range.Any(x => x.Value != null)) {
                for (int colIndex = 0; colIndex < columns; colIndex++) {
                    var col = colIndex + firstCol;
                    var columnModel = columnMapping[colIndex];
                    if (columnModel != null) {
                        var cell = range[row, col]; // note that range[] resets range.Address to equal the new address
                        if (cell.Value == null) {
                            if (!columnModel.Optional)
                                throw new ColumnDataMissingException(columnModel.Name, model.Name);
                        } else {
                            object value;
                            try {
                                if (columnModel.ReadSerializer != null) {
                                    value = columnModel.ReadSerializer(cell);
                                } else {
                                    value = DefaultReadSerializer(cell, columnModel.Type);
                                }
                            } catch (Exception e) {
                                throw new ParseDataException(cell.Address, columnModel.Name, model.Name, e);
                            }
                            if (value != null) {
                                if (columnModel.Member is PropertyInfo propertyInfo) {
                                    propertyInfo.SetMethod.Invoke(obj, new[] { value });
                                } else if (columnModel.Member is FieldInfo fieldInfo) {
                                    fieldInfo.SetValue(obj, value);
                                }
                            }
                        }
                    }
                }
            } else {
                if (model.SkipEmptyRows) {
                    obj = null;
                } else {
                    foreach (var columnModel in columnMapping) {
                        if (!columnModel.Optional)
                            throw new RowEmptyException(model.Name);
                    }
                }
            }
            return obj;
        }

        /// <summary>
        /// Returns an <see cref="ExcelRange"/> mapped to the table of data including headers; defaults to
        /// the entire worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        protected virtual ExcelRange DefaultReadRangeLocator(ExcelWorksheet worksheet)
        {
            var dimension = worksheet.Dimension;
            if (dimension == null)
                return null; // no cells
            return worksheet.Cells[dimension.Start.Row, dimension.Start.Column, dimension.End.Row, dimension.End.Column];
        }

        protected virtual object DefaultCsvReadSerializer(object value, string text, Type dataType, IFormatProvider formatProvider)
        {
            if (value == null) {
                return null;
            }
            if (dataType.IsGenericType && dataType.GetGenericTypeDefinition() == typeof(Nullable<>)) {
                return DefaultCsvReadSerializer(value, text, Nullable.GetUnderlyingType(dataType), formatProvider);
            }
            if (value.GetType() == dataType)
                return value;
            if (dataType == typeof(string))
                return text;
            if (dataType == typeof(DateTime)) {
                if (value is string str)
                    return DateTime.Parse(str, formatProvider, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal);
                return DateTime.FromOADate((double)DefaultCsvReadSerializer(value, text, typeof(double), formatProvider));
            }
            if (dataType == typeof(TimeSpan)) {
                if (value is DateTime dt)
                    return dt.TimeOfDay;
                if (value is string str)
                    try {
                        return TimeSpan.Parse(str, formatProvider);
                    } catch (FormatException) {
                        return DateTime.Parse(str, formatProvider, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal).TimeOfDay;
                    }
            }
            if (dataType == typeof(DateTimeOffset)) {
                throw new NotSupportedException("DateTimeOffset values are not supported");
            }
            if (dataType == typeof(Uri)) {
                return new Uri(text);
            }
            if (dataType == typeof(Guid)) {
                return Guid.Parse(text);
            }
            if (dataType == typeof(bool)) {
                if (value is string str) {
                    switch (str.ToLower()) {
                        case "y":
                        case "1":
                        case "yes":
                        case "true":
                            return true;
                        case "n":
                        case "0":
                        case "no":
                        case "false":
                            return false;
                    }
                }
            }

            if (dataType == typeof(int)) {
                if (value is string str) {
                    var success = int.TryParse(str, NumberStyles.Any, formatProvider, out int result);
                    if (success)
                        return result;

                    return Convert.ToInt32(Math.Floor(double.Parse(str, NumberStyles.Any, formatProvider)));
                }
            }

            if (dataType == typeof(double)) {
                if (value is string str) {
                    return double.Parse(str, NumberStyles.Any, formatProvider);
                }
            }

            if (dataType == typeof(decimal)) {
                if (value is string str) {
                    return decimal.Parse(str, NumberStyles.Any, formatProvider);
                }
            }

            if (dataType == typeof(float)) {
                if (value is string str) {
                    return float.Parse(str, NumberStyles.Any, formatProvider);
                }
            }

            return Convert.ChangeType(value, dataType);
        }

        /// <summary>
        /// Parses the cell and converts it to the requested data type. For nullable types,
        /// it is acceptable to return the underlying type.
        /// <br/><br/>
        /// An exception should be raised if the value cannot be converted.
        /// <br/><br/>
        /// Returning a value of null indicates that the value should be left as the field
        /// default.
        /// </summary>
        protected virtual object DefaultReadSerializer(ExcelRange cell, Type dataType)
        {
            if (cell.Value == null) {
                return null;
            }
            if (dataType.IsGenericType && dataType.GetGenericTypeDefinition() == typeof(Nullable<>)) {
                return DefaultReadSerializer(cell, Nullable.GetUnderlyingType(dataType));
            }
            if (cell.Value.GetType() == dataType)
                return cell.Value;
            if (dataType == typeof(string))
                return cell.Text;
            if (dataType == typeof(DateTime)) {
                if (cell.Value is string str)
                    return DateTime.Parse(str);
                return DateTime.FromOADate((double)DefaultReadSerializer(cell, typeof(double)));
            }
            if (dataType == typeof(TimeSpan)) {
                if (cell.Value is DateTime dt)
                    return dt.TimeOfDay;
                if (cell.Value is string str)
                    return TimeSpan.Parse(str);
                return DateTime.FromOADate((double)DefaultReadSerializer(cell, typeof(double))).TimeOfDay;
            }
            if (dataType == typeof(DateTimeOffset)) {
                throw new NotSupportedException("DateTimeOffset values are not supported");
            }
            if (dataType == typeof(Uri)) {
                return new Uri(cell.Text);
            }
            if (dataType == typeof(Guid)) {
                return Guid.Parse(cell.Text);
            }
            if (dataType == typeof(bool)) {
                if (cell.Value is string str) {
                    switch (str.ToLower()) {
                        case "y":
                        case "yes":
                            return true;
                        case "n":
                        case "no":
                            return false;
                    }
                }
            }
            return Convert.ChangeType(cell.Value, dataType);
        }

        protected virtual void DefaultWriteSerializer(ExcelRange cell, object value)
        {
            /*
            cell.Value = value switch
            {
                null => null,
                DateTime dt => dt.ToOADate(),
                TimeSpan ts => DateTime.FromOADate(0).Add(ts).ToOADate(),
                DateTimeOffset _ => throw new NotSupportedException("DateTimeOffset values are not supported"),
                Guid guid => guid.ToString(),
                Uri uri => uri.ToString(),
                _ => value
            };
            */
            if (value == null)
                cell.Value = null;
            else if (value is DateTime dt)
                cell.Value = dt.ToOADate();
            else if (value is TimeSpan ts)
                cell.Value = DateTime.FromOADate(0).Add(ts).ToOADate();
            else if (value is DateTimeOffset)
                throw new NotSupportedException("DateTimeOffset values are not supported");
            else if (value is Guid guid)
                cell.Value = guid.ToString();
            else if (value is Uri uri)
                cell.Value = uri.ToString();
            else
                cell.Value = value;
        }

        protected virtual void OnWriteFile(ExcelWorkbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            var sheets = GetSheetData();
            for (int i = 0; i < sheets.Count; i++) {
                var sheetModel = Model.Sheets[i];
                var worksheet = workbook.Worksheets.Add(sheetModel.Name);
                OnWriteSheet(worksheet, sheetModel, sheets[i]);
            }
        }

        protected virtual void OnWriteSheet(ExcelWorksheet worksheet, ISheetModel model, IList data)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            if (model == null)
                throw new ArgumentNullException(nameof(model));
            if (data == null)
                throw new ArgumentNullException(nameof(data));
            ExcelRange start = (model.WriteRangeLocator ?? DefaultWriteRangeLocator)(worksheet) ?? throw new InvalidOperationException("No write range specified");
            var headerRow = start.Start.Row;
            var dataRow = headerRow + 1;
            var firstCol = start.Start.Column;
            var columns = model.Columns.Count;
            if (columns == 0)
                return;
            for (int i = 0; i < columns; i++) {
                var columnModel = model.Columns[i];
                var col = firstCol + i;
                var cell = start[headerRow, col]; //note: overwrites start with new address
                cell.Value = columnModel.Name;
                columnModel.HeaderFormatter?.Invoke(cell);
            }
            for (int i = 0; i < data.Count; i++) {
                var cells = start[dataRow + i, firstCol, dataRow + i, firstCol + columns - 1]; //note: overwrites start with new address
                OnWriteRow(cells, model, data[i]);
            }
            if (data.Count > 0) {
                for (int i = 0; i < columns; i++) {
                    var columnModel = model.Columns[i];
                    var col = firstCol + i;
                    var cells = start[dataRow, col, dataRow + data.Count - 1, col]; //note: overwrites start with new address
                    columnModel.ColumnFormatter?.Invoke(cells);
                }
            }
            for (int i = 0; i < columns; i++) {
                var columnModel = model.Columns[i];
                var col = firstCol + i;
                var cells = start[headerRow, col, dataRow + data.Count - 1, col]; //note: overwrites start with new address
                columnModel.WritePolisher?.Invoke(cells);
            }
            var allCells = start[headerRow, firstCol, dataRow + data.Count - 1, firstCol + columns - 1]; //note: overwrites start with new address
            model.WritePolisher?.Invoke(worksheet, allCells);
        }

        protected virtual void OnWriteRow(ExcelRange range, ISheetModel model, object data)
        {
            if (range == null)
                throw new ArgumentNullException(nameof(range));
            if (model == null)
                throw new ArgumentNullException(nameof(model));
            if (data == null)
                throw new ArgumentNullException(nameof(data));
            if (!model.Type.IsAssignableFrom(data.GetType()))
                throw new ArgumentOutOfRangeException("Data type does not match column type");
            var columns = model.Columns.Count;
            var row = range.Start.Row;
            var firstCol = range.Start.Column;
            if (columns != (range.End.Column - range.Start.Column + 1))
                throw new ArgumentOutOfRangeException("Columns in range does not match columns in model");
            if (range.Start.Row != range.End.Row)
                throw new ArgumentOutOfRangeException("Range has more than one row");
            for (int i = 0; i < columns; i++) {
                var cell = range[row, firstCol + i]; //note: overwrites range with new address
                var columnModel = model.Columns[i];
                object value;
                if (columnModel.Member is FieldInfo fieldInfo) {
                    value = fieldInfo.GetValue(data);
                } else if (columnModel.Member is PropertyInfo propertyInfo) {
                    value = propertyInfo.GetValue(data);
                } else {
                    throw new InvalidOperationException("Column member expression is not a field or property");
                }
                var serializer = columnModel.WriteSerializer ?? DefaultWriteSerializer;
                serializer(cell, value);
            }
        }

        protected IReadOnlyList<IList> GetSheetData() => _sheets;

        protected virtual ExcelRange DefaultWriteRangeLocator(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[1, 1];
        }

        public List<T> GetSheet<T>()
        {
            if (!_initialized)
                throw new InvalidOperationException();
            return (List<T>)_sheets[_typeLookup[typeof(T)]];
        }

        public List<T> GetSheet<T>(string name)
        {
            if (!_initialized)
                throw new InvalidOperationException();
            if (name == null)
                throw new ArgumentNullException(nameof(name));
            return (List<T>)_sheets[_sheetNameLookup[name]];
        }

        public virtual ExcelPackage SerializeToExcelPackage()
        {
            var excelPackage = new ExcelPackage();
            excelPackage.Compatibility.IsWorksheets1Based = false;
            OnWriteFile(excelPackage.Workbook);
            return excelPackage;
        }

        public virtual MemoryStream SerializeToStream()
        {
            var stream = new MemoryStream();
            SerializeToStream(stream);
            stream.Position = 0;
            return stream;
        }

        public virtual void SerializeToStream(Stream stream)
        {
            using var excelPackage = SerializeToExcelPackage();
            excelPackage.SaveAs(stream);
        }

        public virtual void SerializeToFile(string filename)
        {
            using var stream = new FileStream(filename, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            SerializeToStream(stream);
        }
    }


}
