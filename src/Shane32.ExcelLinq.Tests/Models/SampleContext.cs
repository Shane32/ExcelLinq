using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml;
using Shane32.ExcelLinq.Builders;
using Shane32.ExcelLinq.Models;

namespace Shane32.ExcelLinq.Tests.Models
{
    public class SampleContext : ExcelContext
    {
        public SampleContext()
        {
        }

        public SampleContext(IExcelModel model) : base(model)
        {
        }

        public IReadOnlyList<IList> TestGetSheetData() => GetSheetData();

        public void TestOnWriteFile(ExcelWorkbook workbook)
        {
            OnWriteFile(workbook);
        }
        public void TestOnWriteSheet(ExcelWorksheet worksheet, ISheetModel model, IList data)
        {
            OnWriteSheet(worksheet, model, data);
        }
        public void TestOnWriteRow(ExcelRange range, ISheetModel model, object data)
        {
            OnWriteRow(range, model, data);
        }
        public object TestOnReadRow(ExcelRange range, ISheetModel model, IColumnModel[] columnMapping)
        {
            return OnReadRow(range, model, columnMapping);
        }

        public IList TestOnReadSheet(ExcelWorksheet worksheet, ISheetModel model)
        {
            return OnReadSheet(worksheet, model);
        }

        public List<IList> TestOnReadFile(ExcelWorkbook workbook)
        {
            return OnReadFile(workbook);
        }

        protected override void OnModelCreating(ExcelModelBuilder modelBuilder)
        {
            var sheet1 = modelBuilder.Sheet<Class1>();
            sheet1.Column(x => x.StringColumn)
                .Optional();
            sheet1.Column(x => x.IntColumn)
                .Optional();
            sheet1.Column(x => x.FloatColumn)
                .Optional();
            sheet1.Column(x => x.DoubleColumn)
                .Optional();
            sheet1.Column(x => x.DateTimeColumn)
                .Optional();
            sheet1.Column(x => x.TimeSpanColumn)
                .Optional();
            sheet1.Column(x => x.BooleanColumn)
                .Optional();
            sheet1.Column(x => x.GuidColumn)
                .Optional();
            sheet1.Column(x => x.UriColumn)
                .Optional();
            sheet1.Column(x => x.DecimalColumn)
                .Optional();
            sheet1.Column(x => x.NullableIntColumn)
                .Optional();
            var sheet2 = modelBuilder.Sheet<Class2>();
            sheet2.Column(x => x.StringColumn);
        }

        public object TestDefaultReadSerializer(ExcelRange cell, Type dataType)
        {
            return DefaultReadSerializer(cell, dataType);
        }

        public void TestDefaultWriteSerializer(ExcelRange cell, object value)
        {
            DefaultWriteSerializer(cell, value);
        }

        public ExcelRange TestDefaultReadRangeLocator(ExcelWorksheet worksheet)
        {
            return DefaultReadRangeLocator(worksheet);
        }
    }
}
