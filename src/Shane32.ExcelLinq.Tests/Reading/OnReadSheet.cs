using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Moq.Protected;
using OfficeOpenXml;
using Shane32.ExcelLinq.Builders;
using Shane32.ExcelLinq.Exceptions;
using Shane32.ExcelLinq.Models;
using Shane32.ExcelLinq.Tests.Models;

namespace Reading
{
    [TestClass]
    public class OnReadSheet
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
            sheet.SetValue(1, 1, "col1");
            sheet.SetValue(2, 1, "test");

            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn, "col1");
            var model = builder.Build();
            var sheetModel = model.Sheets[0];
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadRow", ItExpr.Is<ExcelRange>(x => x.Address == "A2"), sheetModel, ItExpr.Is<IColumnModel[]>(x => x[0] == sheetModel.Columns[0])).Returns(new Class1()).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            var result = context.TestOnReadSheet(sheet, context.Model.Sheets[0]);
            Assert.IsInstanceOfType(result, typeof(List<Class1>));
            var ret = (List<Class1>)result;
            Assert.AreEqual(1, ret.Count);
            mock.Verify();
        }

        [TestMethod]
        public void AltColumnName()
        {
            sheet.SetValue(1, 1, "col1alt");
            sheet.SetValue(2, 1, "test");

            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn, "col1")
                .AlternateName("col1alt");
            var model = builder.Build();
            var sheetModel = model.Sheets[0];
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadRow", ItExpr.Is<ExcelRange>(x => x.Address == "A2"), sheetModel, ItExpr.Is<IColumnModel[]>(x => x[0] == sheetModel.Columns[0])).Returns(new Class1()).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            var result = context.TestOnReadSheet(sheet, context.Model.Sheets[0]);
            Assert.IsInstanceOfType(result, typeof(List<Class1>));
            var ret = (List<Class1>)result;
            Assert.AreEqual(1, ret.Count);
            mock.Verify();
        }

        [TestMethod]
        public void Multiple()
        {
            sheet.SetValue(1, 1, "col1");
            sheet.SetValue(1, 2, "col2");
            sheet.SetValue(2, 1, "test");
            sheet.SetValue(2, 2, 2);
            sheet.SetValue(3, 1, "test2");
            sheet.SetValue(3, 2, 3);

            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.IntColumn, "col2");
            sheetBuilder.Column(x => x.StringColumn, "col1");
            var model = builder.Build();
            var sheetModel = model.Sheets[0];
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadRow", ItExpr.Is<ExcelRange>(x => x.Address == "A2:B2"), sheetModel, ItExpr.Is<IColumnModel[]>(x => x[0] == sheetModel.Columns[1] && x[1] == sheetModel.Columns[0])).Returns(new Class1()).Verifiable();
            mock.Protected().Setup<object>("OnReadRow", ItExpr.Is<ExcelRange>(x => x.Address == "A3:B3"), sheetModel, ItExpr.Is<IColumnModel[]>(x => x[0] == sheetModel.Columns[1] && x[1] == sheetModel.Columns[0])).Returns(new Class1()).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            var result = context.TestOnReadSheet(sheet, context.Model.Sheets[0]);
            Assert.IsInstanceOfType(result, typeof(List<Class1>));
            var ret = (List<Class1>)result;
            Assert.AreEqual(2, ret.Count);
            mock.Verify();
        }

        [TestMethod]
        public void FailRequiredColumns()
        {
            sheet.SetValue(1, 1, "col1");

            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn, "col1");
            sheetBuilder.Column(x => x.IntColumn, "col2");
            var model = builder.Build();
            var sheetModel = model.Sheets[0];
            var context = new SampleContext(model);
            Assert.ThrowsException<ColumnMissingException>(() => {
                var result = context.TestOnReadSheet(sheet, context.Model.Sheets[0]);
            });
        }

        [TestMethod]
        public void FailDuplicateColumns()
        {
            sheet.SetValue(1, 1, "col1");
            sheet.SetValue(1, 2, "col1");

            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn, "col1");
            var model = builder.Build();
            var sheetModel = model.Sheets[0];
            var context = new SampleContext(model);
            Assert.ThrowsException<DuplicateColumnException>(() => {
                var result = context.TestOnReadSheet(sheet, context.Model.Sheets[0]);
            });
        }

        [TestMethod]
        public void FailEmptySheet()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn, "col1");
            var model = builder.Build();
            var sheetModel = model.Sheets[0];
            var context = new SampleContext(model);
            Assert.ThrowsException<SheetEmptyException>(() => {
                var result = context.TestOnReadSheet(sheet, context.Model.Sheets[0]);
            });
        }

        [TestMethod]
        public void EmptySheet()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn, "col1")
                .Optional();
            var model = builder.Build();
            var sheetModel = model.Sheets[0];
            var context = new SampleContext(model);
            var result = context.TestOnReadSheet(sheet, context.Model.Sheets[0]);
            Assert.IsInstanceOfType(result, typeof(List<Class1>));
            var ret = (List<Class1>)result;
            Assert.AreEqual(0, ret.Count);
        }

        [TestMethod]
        public void VerifyAltReadRange()
        {
            sheet.SetValue(2, 2, "col1");
            sheet.SetValue(3, 2, "test");

            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.ReadRangeLocator(worksheet => worksheet.Cells[2, 2, 4, 2]);
            sheetBuilder.Column(x => x.StringColumn, "col1")
                .Optional();
            var model = builder.Build();
            var sheetModel = model.Sheets[0];
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadRow", ItExpr.Is<ExcelRange>(x => x.Address == "B3"), sheetModel, ItExpr.Is<IColumnModel[]>(x => x[0] == sheetModel.Columns[0])).Returns(new Class1()).Verifiable();
            mock.Protected().Setup<object>("OnReadRow", ItExpr.Is<ExcelRange>(x => x.Address == "B4"), sheetModel, ItExpr.Is<IColumnModel[]>(x => x[0] == sheetModel.Columns[0])).Returns(new Class1()).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            var result = context.TestOnReadSheet(sheet, context.Model.Sheets[0]);
            Assert.IsInstanceOfType(result, typeof(List<Class1>));
            var ret = (List<Class1>)result;
            Assert.AreEqual(2, ret.Count);
            mock.Verify();
        }

        [TestMethod]
        public void VerifyAltDefaultReadRange()
        {
            sheet.SetValue(1, 1, "col1");
            sheet.SetValue(2, 1, "test");

            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn, "col1")
                .Optional();
            var model = builder.Build();
            var sheetModel = model.Sheets[0];
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadRow", ItExpr.Is<ExcelRange>(x => x.Address == "A2"), sheetModel, ItExpr.Is<IColumnModel[]>(x => x[0] == sheetModel.Columns[0])).Returns(new Class1()).Verifiable();
            mock.Protected().Setup<object>("OnReadRow", ItExpr.Is<ExcelRange>(x => x.Address == "A3"), sheetModel, ItExpr.Is<IColumnModel[]>(x => x[0] == sheetModel.Columns[0])).Returns(new Class1()).Verifiable();
            mock.Protected().Setup<ExcelRange>("DefaultReadRangeLocator", sheet).Returns(sheet.Cells[1, 1, 3, 1]).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            var result = context.TestOnReadSheet(sheet, context.Model.Sheets[0]);
            Assert.IsInstanceOfType(result, typeof(List<Class1>));
            var ret = (List<Class1>)result;
            Assert.AreEqual(2, ret.Count);
            mock.Verify();
        }

        [TestMethod]
        public void SkipEmptyRows()
        {
            sheet.SetValue(1, 1, "col1");

            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn, "col1");
            var model = builder.Build();
            var sheetModel = model.Sheets[0];
            var mock = new Mock<SampleContext>(model);
            var testret1 = new Class1();
            var testret2 = new Class1();
            mock.Protected().Setup<object>("OnReadRow", ItExpr.Is<ExcelRange>(x => x.Address == "A2"), sheetModel, ItExpr.Is<IColumnModel[]>(x => x[0] == sheetModel.Columns[0])).Returns(testret1).Verifiable();
            mock.Protected().Setup<object>("OnReadRow", ItExpr.Is<ExcelRange>(x => x.Address == "A3"), sheetModel, ItExpr.Is<IColumnModel[]>(x => x[0] == sheetModel.Columns[0])).Returns(null).Verifiable();
            mock.Protected().Setup<object>("OnReadRow", ItExpr.Is<ExcelRange>(x => x.Address == "A4"), sheetModel, ItExpr.Is<IColumnModel[]>(x => x[0] == sheetModel.Columns[0])).Returns(testret2).Verifiable();
            mock.Protected().Setup<ExcelRange>("DefaultReadRangeLocator", sheet).Returns(sheet.Cells[1, 1, 4, 1]).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            var result = context.TestOnReadSheet(sheet, context.Model.Sheets[0]);
            Assert.IsInstanceOfType(result, typeof(List<Class1>));
            var ret = (List<Class1>)result;
            Assert.AreEqual(2, ret.Count);
            Assert.AreEqual(testret1, ret[0]);
            Assert.AreEqual(testret2, ret[1]);
            mock.Verify();
        }

        //[TestMethod]
        //public void AltCreateList()
        //{
        //    sheet.SetValue(1, 1, "col1");
        //    sheet.SetValue(2, 1, "test");

        //    var builder = new ExcelModelBuilder();
        //    var sheetBuilder = builder.Sheet<Class1>();
        //    sheetBuilder.Column(x => x.StringColumn, "col1");
        //    var model = builder.Build();
        //    var sheetModel = model.Sheets[0];
        //    var mock = new Mock<SampleContext>(model);
        //    mock.Protected().Setup<object>("OnReadRow", ItExpr.Is<ExcelRange>(x => x.Address == "A2"), sheetModel, ItExpr.Is<IColumnModel[]>(x => x[0] == sheetModel.Columns[0])).Returns(new Class1()).Verifiable();
        //    mock.Protected().Setup<IList>("CreateListForSheet", typeof(Class1)).Returns(new ListClass1()).Verifiable();
        //    mock.CallBase = true;
        //    var context = mock.Object;
        //    var result = context.TestOnReadSheet(sheet, context.Model.Sheets[0]);
        //    Assert.IsInstanceOfType(result, typeof(ListClass1));
        //    var ret = (ListClass1)result;
        //    Assert.AreEqual(1, ret.Count);
        //    mock.Verify();
        //}

        [TestMethod]
        public void InvalidParametersThrows()
        {
            sheet.SetValue(1, 1, "col1");

            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn, "col1");
            var model = builder.Build();
            var sheetModel = model.Sheets[0];
            var context = new SampleContext(model);
            var resultSuccess = context.TestOnReadSheet(sheet, context.Model.Sheets[0]);
            Assert.ThrowsException<ArgumentNullException>(() => {
                var result = context.TestOnReadSheet(null, context.Model.Sheets[0]);
            });
            Assert.ThrowsException<ArgumentNullException>(() => {
                var result = context.TestOnReadSheet(sheet, null);
            });
        }

        private class ListClass1 : List<Class1> { }
    }
}
