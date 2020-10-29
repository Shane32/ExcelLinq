using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Moq.Protected;
using OfficeOpenXml;
using Shane32.ExcelLinq.Builders;
using Shane32.ExcelLinq.Exceptions;
using Shane32.ExcelLinq.Tests.Models;

namespace Reading
{
    [TestClass]
    public class OnReadRow
    {
        private ExcelPackage package;
        private ExcelWorksheet sheet;

        [TestInitialize]
        public void Initialize()
        {
            package = new ExcelPackage();
            sheet = package.Workbook.Worksheets.Add("Sheet1");
        }

        [TestMethod]
        public void Single()
        {
            sheet.SetValue(1, 1, "test");
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn);
            var context = new SampleContext(builder.Build());
            var result = context.TestOnReadRow(sheet.Cells[1, 1], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[0] });
            Assert.IsInstanceOfType(result, typeof(Class1));
            var ret = (Class1)result;
            Assert.AreEqual("test", ret.StringColumn);
        }

        [TestMethod]
        public void Multiple()
        {
            sheet.SetValue(1, 1, "test1");
            sheet.SetValue(1, 2, 2.5);
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.DoubleColumn);
            sheetBuilder.Column(x => x.StringColumn);
            var context = new SampleContext(builder.Build());
            var result = context.TestOnReadRow(sheet.Cells[1, 1, 1, 2], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[1], context.Model.Sheets[0].Columns[0] });
            Assert.IsInstanceOfType(result, typeof(Class1));
            var ret = (Class1)result;
            Assert.AreEqual("test1", ret.StringColumn);
            Assert.AreEqual(2.5, ret.DoubleColumn);
        }

        [TestMethod]
        public void DoesntCheckMapping()
        {
            sheet.SetValue(1, 1, "test");
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn);
            sheetBuilder.Column(x => x.DoubleColumn);
            var context = new SampleContext(builder.Build());
            var result = context.TestOnReadRow(sheet.Cells[1, 1, 1, 2], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[0], null });
            Assert.IsInstanceOfType(result, typeof(Class1));
            var ret = (Class1)result;
            Assert.AreEqual("test", ret.StringColumn);
        }

        [TestMethod]
        public void DefaultColumns_EmptyRow()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn).Optional();
            sheetBuilder.Column(x => x.DoubleColumn).Optional();
            sheetBuilder.Column(x => x.NullableIntColumn).Optional();
            var context = new SampleContext(builder.Build());
            var result = context.TestOnReadRow(sheet.Cells[1, 1, 1, 3], context.Model.Sheets[0], new[] {
                context.Model.Sheets[0].Columns[0],
                context.Model.Sheets[0].Columns[1],
                context.Model.Sheets[0].Columns[2],
            });
            Assert.IsInstanceOfType(result, typeof(Class1));
            var ret = (Class1)result;
            Assert.AreEqual(null, ret.StringColumn);
            Assert.AreEqual(0, ret.DoubleColumn);
            Assert.AreEqual(null, ret.NullableIntColumn);
        }

        [TestMethod]
        public void DefaultColumns_PartialRow()
        {
            sheet.Cells[1, 4].Value = 4;
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn).Optional();
            sheetBuilder.Column(x => x.DoubleColumn).Optional();
            sheetBuilder.Column(x => x.NullableIntColumn).Optional();
            sheetBuilder.Column(x => x.FloatColumn);
            var context = new SampleContext(builder.Build());
            var result = context.TestOnReadRow(sheet.Cells[1, 1, 1, 4], context.Model.Sheets[0], new[] {
                context.Model.Sheets[0].Columns[0],
                context.Model.Sheets[0].Columns[1],
                context.Model.Sheets[0].Columns[2],
                context.Model.Sheets[0].Columns[3],
            });
            Assert.IsInstanceOfType(result, typeof(Class1));
            var ret = (Class1)result;
            Assert.AreEqual(null, ret.StringColumn);
            Assert.AreEqual(0, ret.DoubleColumn);
            Assert.AreEqual(null, ret.NullableIntColumn);
            Assert.AreEqual(4, ret.FloatColumn);
        }

        [TestMethod]
        public void SkipEmptyRow()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>()
                .SkipEmptyRows();
            sheetBuilder.Column(x => x.StringColumn);
            var context = new SampleContext(builder.Build());
            var result = context.TestOnReadRow(sheet.Cells[1, 1], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[0] });
            Assert.IsNull(result);
        }

        [TestMethod]
        public void InvalidParametersThrows()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>()
                .SkipEmptyRows();
            sheetBuilder.Column(x => x.StringColumn);
            var context = new SampleContext(builder.Build());
            var resultSuccess = context.TestOnReadRow(sheet.Cells[1, 1], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[0] });
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => {
                var result = context.TestOnReadRow(sheet.Cells[1, 1], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[0], null });
            });
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => {
                var result = context.TestOnReadRow(sheet.Cells[1, 1, 1, 2], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[0] });
            });
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => {
                var result = context.TestOnReadRow(sheet.Cells[1, 1, 2, 1], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[0] });
            });
            Assert.ThrowsException<ArgumentNullException>(() => {
                var result = context.TestOnReadRow(null, context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[0] });
            });
            Assert.ThrowsException<ArgumentNullException>(() => {
                var result = context.TestOnReadRow(sheet.Cells[1, 1], null, new[] { context.Model.Sheets[0].Columns[0] });
            });
            Assert.ThrowsException<ArgumentNullException>(() => {
                var result = context.TestOnReadRow(sheet.Cells[1, 1], context.Model.Sheets[0], null);
            });
        }

        [TestMethod]
        public void MissingDataThrows()
        {
            sheet.SetValue(1, 1, "test1");
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.DoubleColumn);
            sheetBuilder.Column(x => x.StringColumn);
            var context = new SampleContext(builder.Build());
            Assert.ThrowsException<ColumnDataMissingException>(() => {
                var result = context.TestOnReadRow(sheet.Cells[1, 1, 1, 2], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[1], context.Model.Sheets[0].Columns[0] });
            });
        }

        [TestMethod]
        public void InvalidDataThrows()
        {
            sheet.SetValue(1, 1, "test1");
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.DoubleColumn);
            var context = new SampleContext(builder.Build());
            Assert.ThrowsException<ParseDataException>(() => {
                var result = context.TestOnReadRow(sheet.Cells[1, 1], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[0] });
            });
        }

        [TestMethod]
        public void EmptyRowThrows()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.DoubleColumn);
            sheetBuilder.Column(x => x.StringColumn);
            var context = new SampleContext(builder.Build());
            Assert.ThrowsException<RowEmptyException>(() => {
                var result = context.TestOnReadRow(sheet.Cells[1, 1, 1, 2], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[1], context.Model.Sheets[0].Columns[0] });
            });
        }

        [TestMethod]
        public void NullableColumnTest()
        {
            sheet.SetValue(1, 1, "test1");
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn);
            sheetBuilder.Column(x => x.NullableIntColumn);
            var context = new SampleContext(builder.Build());
            Assert.ThrowsException<ColumnDataMissingException>(() => {
                var result = context.TestOnReadRow(sheet.Cells[1, 1, 1, 2], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[0], context.Model.Sheets[0].Columns[1] });
            });
        }

        [TestMethod]
        public void AltSerializer()
        {
            sheet.SetValue(1, 1, "test");
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn)
                .ReadSerializer((cell) => "test2");
            var context = new SampleContext(builder.Build());
            var result = context.TestOnReadRow(sheet.Cells[1, 1], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[0] });
            Assert.IsInstanceOfType(result, typeof(Class1));
            var ret = (Class1)result;
            Assert.AreEqual("test2", ret.StringColumn);
        }

        [TestMethod]
        public void VerifyDefaultSerializer()
        {
            sheet.SetValue(1, 1, "test");
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn);
            var mock = new Mock<SampleContext>(builder.Build());
            mock.CallBase = true;
            mock.Protected().Setup<object>("DefaultReadSerializer", ItExpr.Is<ExcelRange>(x => x.Address == "A1"), typeof(string)).Returns("test2").Verifiable();
            var context = mock.Object;
            var result = context.TestOnReadRow(sheet.Cells[1, 1], context.Model.Sheets[0], new[] { context.Model.Sheets[0].Columns[0] });
            Assert.IsInstanceOfType(result, typeof(Class1));
            var ret = (Class1)result;
            Assert.AreEqual("test2", ret.StringColumn);
            mock.Verify();
        }

    }
}
