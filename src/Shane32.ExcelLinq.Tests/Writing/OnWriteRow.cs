using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Moq.Protected;
using OfficeOpenXml;
using Shane32.ExcelLinq.Builders;
using Shane32.ExcelLinq.Tests.Models;

namespace Writing
{
    [TestClass]
    public class OnWriteRow
    {
        private ExcelPackage package;
        private ExcelWorksheet sheet;
        private SampleContext context;
        private Class1 testRow1;
        private Class1 testRow2;

        [TestInitialize]
        public void Initialize()
        {
            package = new ExcelPackage();
            sheet = package.Workbook.Worksheets.Add("Sheet1");
            context = new SampleContext();
            testRow1 = new Class1 {
                BooleanColumn = true,
                DateTimeColumn = DateTime.Parse("5/22/2020 1:30 PM"),
                DecimalColumn = 55.9m,
                DoubleColumn = 3.333,
                FloatColumn = 2.222f,
                GuidColumn = Guid.NewGuid(),
                IntColumn = 1,
                NullableIntColumn = 92,
                StringColumn = "test",
                TimeSpanColumn = TimeSpan.FromMinutes(125.5),
                UriColumn = new Uri("http://localhost/uri")
            };
            testRow2 = new Class1 {
                BooleanColumn = false,
                DateTimeColumn = DateTime.Parse("5/22/2020 3:30 PM"),
                DecimalColumn = 155.9m,
                DoubleColumn = 13.333,
                FloatColumn = 12.222f,
                GuidColumn = Guid.NewGuid(),
                IntColumn = 11,
                NullableIntColumn = 5,
                StringColumn = "test2",
                TimeSpanColumn = TimeSpan.FromHours(14),
                UriColumn = new Uri("http://localhost/uri2")
            };
        }

        [TestMethod]
        public void Simple()
        {
            context.TestOnWriteRow(sheet.Cells[2, 2, 2, 12], context.Model.Sheets[0], testRow1);
        }

        [TestMethod]
        public void Simple_CheckValues()
        {
            context.TestOnWriteRow(sheet.Cells[2, 2, 2, 12], context.Model.Sheets[0], testRow1);
            context.TestOnWriteRow(sheet.Cells[3, 2, 3, 12], context.Model.Sheets[0], testRow2);

            Assert.AreEqual(testRow1.StringColumn, sheet.Cells[2, 2].Value);
            Assert.AreEqual(testRow1.IntColumn, sheet.Cells[2, 3].Value);
            Assert.AreEqual(testRow1.FloatColumn, sheet.Cells[2, 4].Value);
            Assert.AreEqual(testRow1.DoubleColumn, sheet.Cells[2, 5].Value);
            Assert.AreEqual(testRow1.DateTimeColumn.ToOADate(), sheet.Cells[2, 6].Value);
            Assert.AreEqual(DateTime.FromOADate(0).Add(testRow1.TimeSpanColumn).ToOADate(), sheet.Cells[2, 7].Value);
            Assert.AreEqual(testRow1.BooleanColumn, sheet.Cells[2, 8].Value);
            Assert.AreEqual(testRow1.GuidColumn.ToString(), sheet.Cells[2, 9].Value);
            Assert.AreEqual(testRow1.UriColumn.ToString(), sheet.Cells[2, 10].Value);
            Assert.AreEqual(testRow1.DecimalColumn, sheet.Cells[2, 11].Value);
            Assert.AreEqual(testRow1.NullableIntColumn, sheet.Cells[2, 12].Value);

            Assert.AreEqual(testRow2.StringColumn, sheet.Cells[3, 2].Value);
            Assert.AreEqual(testRow2.IntColumn, sheet.Cells[3, 3].Value);
            Assert.AreEqual(testRow2.FloatColumn, sheet.Cells[3, 4].Value);
            Assert.AreEqual(testRow2.DoubleColumn, sheet.Cells[3, 5].Value);
            Assert.AreEqual(testRow2.DateTimeColumn.ToOADate(), sheet.Cells[3, 6].Value);
            Assert.AreEqual(DateTime.FromOADate(0).Add(testRow2.TimeSpanColumn).ToOADate(), sheet.Cells[3, 7].Value);
            Assert.AreEqual(testRow2.BooleanColumn, sheet.Cells[3, 8].Value);
            Assert.AreEqual(testRow2.GuidColumn.ToString(), sheet.Cells[3, 9].Value);
            Assert.AreEqual(testRow2.UriColumn.ToString(), sheet.Cells[3, 10].Value);
            Assert.AreEqual(testRow2.DecimalColumn, sheet.Cells[3, 11].Value);
            Assert.AreEqual(testRow2.NullableIntColumn, sheet.Cells[3, 12].Value);
        }

        [TestMethod]
        public void SpecifyWriteSerializer()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn)
                .WriteSerializer((range, value) => range.Value = "SpecifyWriteSerializer")
                .Optional();
            var context = new SampleContext(builder.Build());

            context.TestOnWriteRow(sheet.Cells[1, 1], context.Model.Sheets[0], testRow1);
            Assert.AreEqual("SpecifyWriteSerializer", sheet.Cells[1, 1].Value);
        }

        [TestMethod]
        public void DefaultWriteSerializer()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn)
                .Optional();
            var mock = new Mock<SampleContext>(builder.Build());
            mock.CallBase = true;
            mock.Protected().Setup("DefaultWriteSerializer", ItExpr.IsAny<ExcelRange>(), testRow1.StringColumn)
                .Callback<ExcelRange, object>((range, value) => range.Value = "SpecifyWriteSerializer2")
                .Verifiable();
            var context = mock.Object;

            context.TestOnWriteRow(sheet.Cells[1, 1], context.Model.Sheets[0], testRow1);
            Assert.AreEqual("SpecifyWriteSerializer2", sheet.Cells[1, 1].Value);
            mock.Verify();
        }

        [TestMethod]
        public void InvalidParametersThrows()
        {
            context.TestOnWriteRow(sheet.Cells[1, 1, 1, 11], context.Model.Sheets[0], testRow1);
            Assert.ThrowsException<ArgumentNullException>(() => {
                context.TestOnWriteRow(null, context.Model.Sheets[0], testRow1);
            });
            Assert.ThrowsException<ArgumentNullException>(() => {
                context.TestOnWriteRow(sheet.Cells[1, 1, 1, 11], null, testRow1);
            });
            Assert.ThrowsException<ArgumentNullException>(() => {
                context.TestOnWriteRow(sheet.Cells[1, 1, 1, 11], context.Model.Sheets[0], null);
            });
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => {
                context.TestOnWriteRow(sheet.Cells[1, 1, 1, 11], context.Model.Sheets[0], new Class2());
            });
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => {
                context.TestOnWriteRow(sheet.Cells[1, 1, 1, 10], context.Model.Sheets[0], testRow1);
            });
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => {
                context.TestOnWriteRow(sheet.Cells[1, 1, 2, 11], context.Model.Sheets[0], testRow1);
            });
        }

        
    }
}
