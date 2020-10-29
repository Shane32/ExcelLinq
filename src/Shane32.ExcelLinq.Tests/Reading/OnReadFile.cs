using System;
using System.Collections.Generic;
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
    public class OnReadFile
    {
        private ExcelPackage package;
        private ExcelWorkbook workbook;

        [TestInitialize]
        public void Initialize()
        {
            package = new ExcelPackage();
            workbook = package.Workbook;
        }

        [TestMethod]
        public void Single()
        {
            var sheet = workbook.Worksheets.Add("Sheet1");
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class1>("Sheet1");
            var model = builder.Build();
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadSheet", sheet, model.Sheets[0]).Returns(new List<Class1>()).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            var result = context.TestOnReadFile(workbook);
            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count);
            Assert.IsInstanceOfType(result[0], typeof(List<Class1>));
            mock.Verify();
        }

        [TestMethod]
        public void SingleAltName()
        {
            var sheet = workbook.Worksheets.Add("Sheet1");
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class1>()
                .AlternateName("Sheet1");
            var model = builder.Build();
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadSheet", sheet, model.Sheets[0]).Returns(new List<Class1>()).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            var result = context.TestOnReadFile(workbook);
            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count);
            Assert.IsInstanceOfType(result[0], typeof(List<Class1>));
            mock.Verify();
        }

        [TestMethod]
        public void Multiple()
        {
            var sheet = workbook.Worksheets.Add("Sheet1");
            var sheet2 = workbook.Worksheets.Add("Sheet2");
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class2>("Sheet2");
            builder.Sheet<Class1>("Sheet1");
            var model = builder.Build();
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadSheet", sheet, model.Sheets[1]).Returns(new List<Class1>()).Verifiable();
            mock.Protected().Setup<object>("OnReadSheet", sheet2, model.Sheets[0]).Returns(new List<Class2>()).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            var result = context.TestOnReadFile(workbook);
            Assert.IsNotNull(result);
            Assert.AreEqual(2, result.Count);
            Assert.IsInstanceOfType(result[0], typeof(List<Class2>));
            Assert.IsInstanceOfType(result[1], typeof(List<Class1>));
            mock.Verify();
        }

        [TestMethod]
        public void MultipleIgnoreNames()
        {
            var sheet = workbook.Worksheets.Add("Sheet1");
            var sheet2 = workbook.Worksheets.Add("Sheet2");
            var builder = new ExcelModelBuilder();
            builder.IgnoreSheetNames();
            builder.Sheet<Class1>();
            builder.Sheet<Class2>();
            var model = builder.Build();
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadSheet", sheet, model.Sheets[0]).Returns(new List<Class1>()).Verifiable();
            mock.Protected().Setup<object>("OnReadSheet", sheet2, model.Sheets[1]).Returns(new List<Class2>()).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            var result = context.TestOnReadFile(workbook);
            Assert.IsNotNull(result);
            Assert.AreEqual(2, result.Count);
            Assert.IsInstanceOfType(result[0], typeof(List<Class1>));
            Assert.IsInstanceOfType(result[1], typeof(List<Class2>));
            mock.Verify();
        }

        [TestMethod]
        public void OptionalSheet()
        {
            var sheet = workbook.Worksheets.Add("Sheet1");
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class1>().Optional();
            builder.Sheet<Class2>("Sheet1");
            var model = builder.Build();
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadSheet", sheet, model.Sheets[1]).Returns(new List<Class2>()).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            var result = context.TestOnReadFile(workbook);
            Assert.IsNotNull(result);
            Assert.AreEqual(2, result.Count);
            Assert.IsInstanceOfType(result[0], typeof(List<Class1>));
            Assert.IsInstanceOfType(result[1], typeof(List<Class2>));
            mock.Verify();
        }

        [TestMethod]
        public void ExtraSheet()
        {
            var sheet = workbook.Worksheets.Add("Sheet1");
            var sheet2 = workbook.Worksheets.Add("Sheet2");
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class1>("Sheet2");
            var model = builder.Build();
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadSheet", sheet2, model.Sheets[0]).Returns(new List<Class1>()).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            var result = context.TestOnReadFile(workbook);
            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count);
            Assert.IsInstanceOfType(result[0], typeof(List<Class1>));
            mock.Verify();
        }

        [TestMethod]
        public void NullResultFromOnReadSheet_Named()
        {
            var sheet = workbook.Worksheets.Add("Sheet1");
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class1>("Sheet1");
            var model = builder.Build();
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadSheet", sheet, model.Sheets[0]).Returns(null).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            Assert.ThrowsException<InvalidOperationException>(() => {
                var result = context.TestOnReadFile(workbook);
            });
        }

        [TestMethod]
        public void NullResultFromOnReadSheet_Unnamed()
        {
            var sheet = workbook.Worksheets.Add("Sheet1");
            var builder = new ExcelModelBuilder();
            builder.IgnoreSheetNames();
            builder.Sheet<Class1>();
            var model = builder.Build();
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadSheet", sheet, model.Sheets[0]).Returns(null).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            Assert.ThrowsException<InvalidOperationException>(() => {
                var result = context.TestOnReadFile(workbook);
            });
        }

        [TestMethod]
        public void DuplicateSheet()
        {
            var sheet = workbook.Worksheets.Add("Sheet1");
            var sheet2 = workbook.Worksheets.Add("Sheet2");
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class1>("Sheet1").AlternateName("Sheet2");
            var model = builder.Build();
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadSheet", sheet, model.Sheets[0]).Returns(new List<Class1>()).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            Assert.ThrowsException<DuplicateSheetException>(() => {
                var result = context.TestOnReadFile(workbook);
            });
        }

        [TestMethod]
        public void MissingSheet()
        {
            var sheet = workbook.Worksheets.Add("Sheet1");
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class1>();
            var model = builder.Build();
            var mock = new Mock<SampleContext>(model);
            mock.Protected().Setup<object>("OnReadSheet", sheet, model.Sheets[0]).Returns(new List<Class1>()).Verifiable();
            mock.CallBase = true;
            var context = mock.Object;
            Assert.ThrowsException<SheetMissingException>(() => {
                var result = context.TestOnReadFile(workbook);
            });
        }

        [TestMethod]
        public void InvalidParametersThrows()
        {
            var context = new SampleContext();
            Assert.ThrowsException<ArgumentNullException>(() => {
                context.TestOnReadFile(null);
            });
        }
    }
}
