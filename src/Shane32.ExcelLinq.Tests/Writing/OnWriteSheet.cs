using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Moq.Protected;
using OfficeOpenXml;
using Shane32.ExcelLinq.Builders;
using Shane32.ExcelLinq.Models;
using Shane32.ExcelLinq.Tests.Models;

namespace Writing
{
    [TestClass]
    public class OnWriteSheet
    {
        private ExcelPackage package;
        private ExcelWorksheet sheet;

        [TestInitialize]
        public void Initialize()
        {
            package = new ExcelPackage();
            sheet = package.Workbook.Worksheets.Add("Shet1");
        }

        [TestMethod]
        public void Simple()
        {
            var mock = new Mock<SampleContext>();
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(builder => {
                var sheetBuilder = builder.Sheet<Class1>();
                sheetBuilder.Column(x => x.StringColumn);
            }).Verifiable();
            var row1 = new Class1();
            mock.Protected().Setup("OnWriteRow", ItExpr.Is<ExcelRange>(x => x.Address == "A2"), ItExpr.IsAny<ISheetModel>(), row1)
                .Callback<ExcelRange, ISheetModel, object>((a, b, c) => { })
                .Verifiable();
            var context = mock.Object;
            var data = new List<Class1>() { row1 };
            context.TestOnWriteSheet(sheet, context.Model.Sheets[0], data);
            mock.Verify();
        }

        [TestMethod]
        public void VerifyHeaders()
        {
            var mock = new Mock<SampleContext>();
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(builder => {
                var sheetBuilder = builder.Sheet<Class1>();
                sheetBuilder.WriteRangeLocator(worksheet => worksheet.Cells[2, 2]);
                sheetBuilder.Column(x => x.StringColumn).AlternateName("col1");
                sheetBuilder.Column(x => x.IntColumn, "col2").AlternateName("col3");
            }).Verifiable();
            var context = mock.Object;
            var data = new List<Class1>() { };
            context.TestOnWriteSheet(sheet, context.Model.Sheets[0], data);
            mock.Verify();
            Assert.AreEqual("StringColumn", sheet.Cells[2, 2].Value);
            Assert.AreEqual("col2", sheet.Cells[2, 3].Value);
        }

        [TestMethod]
        public void Multiple()
        {
            var mock = new Mock<SampleContext>();
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(builder => {
                var sheetBuilder = builder.Sheet<Class1>();
                sheetBuilder.Column(x => x.StringColumn);
                sheetBuilder.Column(x => x.IntColumn);
            }).Verifiable();
            var row1 = new Class1();
            var row2 = new Class1();
            mock.Protected().Setup("OnWriteRow", ItExpr.Is<ExcelRange>(x => x.Address == "A2:B2"), ItExpr.IsAny<ISheetModel>(), row1)
                .Callback<ExcelRange, ISheetModel, object>((a, b, c) => { })
                .Verifiable();
            mock.Protected().Setup("OnWriteRow", ItExpr.Is<ExcelRange>(x => x.Address == "A3:B3"), ItExpr.IsAny<ISheetModel>(), row2)
                .Callback<ExcelRange, ISheetModel, object>((a, b, c) => { })
                .Verifiable();
            var context = mock.Object;
            var data = new List<Class1>() { row1, row2 };
            context.TestOnWriteSheet(sheet, context.Model.Sheets[0], data);
            mock.Verify();
        }

        [TestMethod]
        public void RowsWithNoColumns()
        {
            var mock = new Mock<SampleContext>();
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(builder => {
                builder.Sheet<Class1>();
            }).Verifiable();
            var row1 = new Class1();
            mock.Protected().Setup("OnWriteRow", ItExpr.IsAny<ExcelRange>(), ItExpr.IsAny<ISheetModel>(), row1)
                .Throws<Exception>();
            var context = mock.Object;
            context.GetSheet<Class1>().Add(row1);
            var data = new List<Class1>() { row1 };
            context.TestOnWriteSheet(sheet, context.Model.Sheets[0], data);
            mock.Verify();
        }

        [TestMethod]
        public void DefaultWriteRangeLocator()
        {
            var mock = new Mock<SampleContext>();
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>())
                .Callback<ExcelModelBuilder>(builder => {
                    builder.Sheet<Class1>()
                        .Column(x => x.StringColumn);
                })
                .Verifiable();
            mock.Protected().Setup<ExcelRange>("DefaultWriteRangeLocator", sheet)
                .Returns<ExcelWorksheet>(worksheet => worksheet.Cells[2, 2])
                .Verifiable();
            var row1 = new Class1();
            mock.Protected().Setup("OnWriteRow", ItExpr.Is<ExcelRange>(x => x.Address == "B3"), ItExpr.IsAny<ISheetModel>(), row1)
                .Callback<ExcelRange, ISheetModel, object>((a, b, c) => { })
                .Verifiable();
            var context = mock.Object;
            var data = new List<Class1>() { row1 };
            context.TestOnWriteSheet(sheet, context.Model.Sheets[0], data);
            mock.Verify();
        }

        [TestMethod]
        public void DefaultWriteRangeLocator_InvalidThrows()
        {
            var mock = new Mock<SampleContext>();
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>())
                .Callback<ExcelModelBuilder>(builder => {
                    builder.Sheet<Class1>()
                        .Column(x => x.StringColumn);
                })
                .Verifiable();
            mock.Protected().Setup<ExcelRange>("DefaultWriteRangeLocator", sheet)
                .Returns<ExcelWorksheet>(worksheet => null)
                .Verifiable();
            var row1 = new Class1();
            var context = mock.Object;
            var data = new List<Class1>() { row1 };
            Assert.ThrowsException<InvalidOperationException>(() => {
                context.TestOnWriteSheet(sheet, context.Model.Sheets[0], data);
            });
            mock.Verify();
        }

        [TestMethod]
        public void TestSheetAndColumnProps()
        {
            var mockWrl = new Mock<Func<ExcelWorksheet, ExcelRange>>();
            mockWrl.Setup(f => f(sheet)).Returns(sheet.Cells[2, 2]).Verifiable();
            var mockWp = new Mock<Action<ExcelWorksheet, ExcelRange>>();
            mockWp.Setup(f => f(sheet, It.Is<ExcelRange>(x => x.Address == "B2:C4"))).Verifiable();

            var mockHf = new Mock<Action<ExcelRange>>();
            mockHf.Setup(f => f(It.Is<ExcelRange>(x => x.Address == "B2"))).Verifiable();
            var mockCf = new Mock<Action<ExcelRange>>();
            mockCf.Setup(f => f(It.Is<ExcelRange>(x => x.Address == "B3:B4"))).Verifiable();
            var mockCwp = new Mock<Action<ExcelRange>>();
            mockCwp.Setup(f => f(It.Is<ExcelRange>(x => x.Address == "B2:B4"))).Verifiable();

            var mockHf2 = new Mock<Action<ExcelRange>>();
            mockHf2.Setup(f => f(It.Is<ExcelRange>(x => x.Address == "C2"))).Verifiable();
            var mockCf2 = new Mock<Action<ExcelRange>>();
            mockCf2.Setup(f => f(It.Is<ExcelRange>(x => x.Address == "C3:C4"))).Verifiable();
            var mockCwp2 = new Mock<Action<ExcelRange>>();
            mockCwp2.Setup(f => f(It.Is<ExcelRange>(x => x.Address == "C2:C4"))).Verifiable();

            var mock = new Mock<SampleContext>();
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(builder => {
                var sheetBuilder = builder.Sheet<Class1>()
                    .WriteRangeLocator(mockWrl.Object)
                    .WritePolisher(mockWp.Object);
                sheetBuilder.Column(x => x.StringColumn)
                    .ColumnFormatter(mockCf.Object)
                    .HeaderFormatter(mockHf.Object)
                    .WritePolisher(mockCwp.Object);
                sheetBuilder.Column(x => x.IntColumn)
                    .ColumnFormatter(mockCf2.Object)
                    .HeaderFormatter(mockHf2.Object)
                    .WritePolisher(mockCwp2.Object);
            }).Verifiable();
            var row1 = new Class1();
            var row2 = new Class1();
            mock.Protected().Setup("OnWriteRow", ItExpr.Is<ExcelRange>(x => x.Address == "B3:C3"), ItExpr.IsAny<ISheetModel>(), row1)
                .Callback<ExcelRange, ISheetModel, object>((a, b, c) => { })
                .Verifiable();
            mock.Protected().Setup("OnWriteRow", ItExpr.Is<ExcelRange>(x => x.Address == "B4:C4"), ItExpr.IsAny<ISheetModel>(), row2)
                .Callback<ExcelRange, ISheetModel, object>((a, b, c) => { })
                .Verifiable();
            var context = mock.Object;
            var data = new List<Class1>() { row1, row2 };
            context.TestOnWriteSheet(sheet, context.Model.Sheets[0], data);

            mock.Verify();
            mockWp.Verify();

            mockCf.Verify();
            mockHf.Verify();
            mockCwp.Verify();

            mockCf2.Verify();
            mockHf2.Verify();
            mockCwp2.Verify();
        }

        [TestMethod]
        public void InvalidParametersThrows()
        {
            var mock = new Mock<SampleContext>();
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(builder => {
                var sheetBuilder = builder.Sheet<Class1>();
                sheetBuilder.Column(x => x.StringColumn);
            });
            var row1 = new Class1();
            mock.Protected().Setup("OnWriteRow", ItExpr.Is<ExcelRange>(x => x.Address == "A2"), ItExpr.IsAny<ISheetModel>(), row1)
                .Callback<ExcelRange, ISheetModel, object>((a, b, c) => { });
            var context = mock.Object;
            var data = new List<Class1>() { row1 };
            context.TestOnWriteSheet(sheet, context.Model.Sheets[0], data);

            Assert.ThrowsException<ArgumentNullException>(() => {
                context.TestOnWriteSheet(null, context.Model.Sheets[0], data);
            });
            Assert.ThrowsException<ArgumentNullException>(() => {
                context.TestOnWriteSheet(sheet, null, data);
            });
            Assert.ThrowsException<ArgumentNullException>(() => {
                context.TestOnWriteSheet(sheet, context.Model.Sheets[0], null);
            });
        }
    }
}
