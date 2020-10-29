using System;
using System.Collections;
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
    public class OnWriteFile
    {
        [TestMethod]
        public void Multiple()
        {
            var package = new ExcelPackage();
            var workbook = package.Workbook;
            var mock = new Mock<SampleContext>();
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(builder => {
                builder.Sheet<Class1>();
                builder.Sheet<Class2>("Sheet2");
            }).Verifiable();
            var context = mock.Object;
            var model = context.Model;
            var data = context.TestGetSheetData();
            Assert.AreEqual(2, data.Count);
            mock.Protected().Setup("OnWriteSheet", ItExpr.Is<ExcelWorksheet>(x => x.Name == "Class1"), model.Sheets[0], data[0])
                .Callback<ExcelWorksheet, ISheetModel, IList>((a, b, c) => { })
                .Verifiable();
            mock.Protected().Setup("OnWriteSheet", ItExpr.Is<ExcelWorksheet>(x => x.Name == "Sheet2"), model.Sheets[1], data[1])
                .Callback<ExcelWorksheet, ISheetModel, IList>((a, b, c) => { })
                .Verifiable();
            context.TestOnWriteFile(workbook);
            mock.Verify();
            Assert.AreEqual(2, workbook.Worksheets.Count);
            Assert.AreEqual("Class1", workbook.Worksheets[0].Name);
            Assert.AreEqual("Sheet2", workbook.Worksheets[1].Name);
        }

        [TestMethod]
        public void None()
        {
            var model = new ExcelModelBuilder().Build();
            var context = new SampleContext(model);
            var workbook = new ExcelPackage().Workbook;
            var data = context.TestGetSheetData();
            Assert.AreEqual(0, data.Count);
            context.TestOnWriteFile(workbook);
            Assert.AreEqual(0, workbook.Worksheets.Count);
        }

        [TestMethod]
        public void InvalidParametersThrows()
        {
            var context = new SampleContext();
            Assert.ThrowsException<ArgumentNullException>(() => context.TestOnWriteFile(null));
        }
    }
}
