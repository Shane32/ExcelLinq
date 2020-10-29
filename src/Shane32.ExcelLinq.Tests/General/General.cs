using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Moq.Protected;
using OfficeOpenXml;
using Shane32.ExcelLinq;
using Shane32.ExcelLinq.Builders;
using Shane32.ExcelLinq.Tests.Models;

namespace General
{
    [TestClass]
    public class General
    {
        private ExcelWorksheet sheet;
        private ExcelRange cell;
        private SampleContext context;

        [TestInitialize]
        public void InitializeCell()
        {
            var package = new ExcelPackage();
            sheet = package.Workbook.Worksheets.Add("Sheet1");
            cell = sheet.Cells[1, 1];
            context = new SampleContext();
        }

        [TestMethod]
        public void SelectAll()
        {
            sheet.Cells[1, 1].Value = "test1";
            sheet.Cells[2, 3].Value = "test2";
            Assert.AreEqual(1, sheet.Dimension.Start.Row);
            Assert.AreEqual(1, sheet.Dimension.Start.Column);
            Assert.AreEqual(2, sheet.Dimension.End.Row);
            Assert.AreEqual(3, sheet.Dimension.End.Column);
        }

        [TestMethod]
        public void DefaultReadRange()
        {
            sheet.Cells[1, 1].Value = "test1";
            sheet.Cells[2, 3].Value = "test2";
            var range = context.TestDefaultReadRangeLocator(sheet);
            Assert.AreEqual(1, range.Start.Row);
            Assert.AreEqual(1, range.Start.Column);
            Assert.AreEqual(2, range.End.Row);
            Assert.AreEqual(3, range.End.Column);
        }

        [TestMethod]
        public void HomeIs11()
        {
            Assert.AreEqual(1, cell["A1"].Start.Row);
            Assert.AreEqual(1, cell["A1"].Start.Column);
        }

        [TestMethod]
        public void Constructor_Default()
        {
            var mock = new Mock<ExcelContext>();
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(
                (builder) => {
                    builder.Sheet<Class1>();
                }).Verifiable();
            var context = mock.Object;
            Assert.IsInstanceOfType(context.GetSheet<Class1>(), typeof(List<Class1>));
            Assert.AreEqual(0, context.GetSheet<Class1>().Count);
            Assert.ThrowsException<KeyNotFoundException>(() => {
                context.GetSheet<Class2>();
            });
            mock.Verify();
        }

        //[TestMethod]
        //public void Constructor_Model()
        //{
        //    var builder = new ExcelModelBuilder();
        //    builder.Sheet<Class1>();
        //    var model = builder.Build();
        //    var mock = new Mock<ExcelContext>(model);
        //    mock.CallBase = true;
        //    mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Throws<Exception>();
        //    var context = mock.Object;
        //    Assert.IsInstanceOfType(context.GetSheet<Class1>(), typeof(List<Class1>));
        //    Assert.AreEqual(0, context.GetSheet<Class1>().Count);
        //    Assert.ThrowsException<KeyNotFoundException>(() => {
        //        context.GetSheet<Class2>();
        //    });
        //    mock.Verify();
        //}

        [TestMethod]
        public void Constructor_Stream()
        {
            var package = new ExcelPackage();
            var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
            var stream = new System.IO.MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            var mock = new Mock<ExcelContext>(stream);
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(
                (builder) => {
                    builder.Sheet<Class1>("Sheet1");
                }).Verifiable();
            var class1Data = new List<Class1>();
            mock.Protected().Setup<List<IList>>("OnReadFile", ItExpr.IsAny<ExcelWorkbook>()).Returns<ExcelWorkbook>(
                workbook => {
                    return new List<IList>(new IList[] { class1Data });
                }).Verifiable();
            var context = mock.Object;
            Assert.AreEqual(class1Data, context.GetSheet<Class1>());
            Assert.ThrowsException<KeyNotFoundException>(() => {
                context.GetSheet<Class2>();
            });
            mock.Verify();
        }

        //[TestMethod]
        //public void Constructor_ModelStream()
        //{
        //    var package = new ExcelPackage();
        //    var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
        //    var stream = new System.IO.MemoryStream();
        //    package.SaveAs(stream);
        //    stream.Position = 0;

        //    var builder = new ExcelModelBuilder();
        //    builder.Sheet<Class1>("Sheet1");
        //    var model = builder.Build();
        //    var mock = new Mock<ExcelContext>(model, stream);
        //    mock.CallBase = true;
        //    mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Throws<Exception>();
        //    var class1Data = new List<Class1>();
        //    mock.Protected().Setup<List<IList>>("OnReadFile", ItExpr.IsAny<ExcelWorkbook>()).Returns<ExcelWorkbook>(
        //        workbook => {
        //            return new List<IList>(new IList[] { class1Data });
        //        }).Verifiable();
        //    var context = mock.Object;
        //    Assert.AreEqual(class1Data, context.GetSheet<Class1>());
        //    Assert.ThrowsException<KeyNotFoundException>(() => {
        //        context.GetSheet<Class2>();
        //    });
        //    mock.Verify();
        //}

        [TestMethod]
        public void Constructor_Stream_InvalidOnReadFile_1()
        {
            var package = new ExcelPackage();
            var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
            var stream = new System.IO.MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            var mock = new Mock<ExcelContext>(stream);
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(
                (builder) => {
                    builder.Sheet<Class1>("Sheet1");
                }).Verifiable();
            mock.Protected().Setup<List<IList>>("OnReadFile", ItExpr.IsAny<ExcelWorkbook>()).Returns<ExcelWorkbook>(
                workbook => {
                    return null;
                }).Verifiable();
            AssertThrowsInnerException<InvalidOperationException>(() => {
                var context = mock.Object;
            });
        }

        [TestMethod]
        public void Constructor_Stream_InvalidOnReadFile_2()
        {
            var package = new ExcelPackage();
            var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
            var stream = new System.IO.MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            var mock = new Mock<ExcelContext>(stream);
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(
                (builder) => {
                    builder.Sheet<Class1>("Sheet1");
                }).Verifiable();
            mock.Protected().Setup<List<IList>>("OnReadFile", ItExpr.IsAny<ExcelWorkbook>()).Returns<ExcelWorkbook>(
                workbook => {
                    return new List<IList>(); //wrong number of sheets
                }).Verifiable();
            AssertThrowsInnerException<InvalidOperationException>(() => {
                var context = mock.Object;
            });
        }

        [TestMethod]
        public void Constructor_Stream_InvalidOnReadFile_3()
        {
            var package = new ExcelPackage();
            var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
            var stream = new System.IO.MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            var mock = new Mock<ExcelContext>(stream);
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(
                (builder) => {
                    builder.Sheet<Class1>("Sheet1");
                }).Verifiable();
            mock.Protected().Setup<List<IList>>("OnReadFile", ItExpr.IsAny<ExcelWorkbook>()).Returns<ExcelWorkbook>(
                workbook => {
                    return new List<IList>(new IList[] { null }); //sheet is null
                }).Verifiable();
            AssertThrowsInnerException<InvalidOperationException>(() => {
                var context = mock.Object;
            });
        }

        [TestMethod]
        public void Constructor_Stream_InvalidOnReadFile_4()
        {
            var package = new ExcelPackage();
            var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
            var stream = new System.IO.MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            var mock = new Mock<ExcelContext>(stream);
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(
                (builder) => {
                    builder.Sheet<Class1>("Sheet1");
                }).Verifiable();
            mock.Protected().Setup<List<IList>>("OnReadFile", ItExpr.IsAny<ExcelWorkbook>()).Returns<ExcelWorkbook>(
                workbook => {
                    return new List<IList>(new IList[] { new List<Class2>() }); //wrong type of sheet
                }).Verifiable();
            AssertThrowsInnerException<InvalidOperationException>(() => {
                var context = mock.Object;
            });
        }

        private void AssertThrowsInnerException<T>(Action action) where T : Exception
        {
            try {
                action();
                Assert.Fail("No exception was thrown");
            }
            catch (TargetInvocationException e) {
                Assert.IsInstanceOfType(e.InnerException, typeof(T));
            }
        }
        
    }
}
