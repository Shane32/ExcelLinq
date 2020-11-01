using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using Shane32.ExcelLinq.Tests.Models;

namespace Serializers
{
    [TestClass]
    public class DefaultWriteSerializer
    {
        private ExcelRange cell;
        private SampleContext context;

        [TestInitialize]
        public void InitializeCell()
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");
            cell = worksheet.Cells[1, 1];
            context = new SampleContext();
        }

        [DataTestMethod]
        [DataRow("test", "test", DisplayName = "string")]
        [DataRow(1, 1, DisplayName = "int")]
        [DataRow(2.1f, 2.1f, DisplayName = "float")]
        [DataRow(2.2, 2.2, DisplayName = "double")]
        [DataRow(true, true, DisplayName = "bool")]
        [DataRow(null, null, DisplayName = "null")]
        public void Test(object input, object expectedOutput)
        {
            context.TestDefaultWriteSerializer(cell, input);
            Assert.AreEqual(expectedOutput, cell.Value);
        }

        [TestMethod]
        public void DateTimeTest()
        {
            var dateTime = DateTime.FromOADate(44137.54);
            context.TestDefaultWriteSerializer(cell, dateTime);
            Assert.AreEqual(44137.54, cell.Value);
        }

        [TestMethod]
        public void DateTimeOffsetTest()
        {
            var dateTime = DateTimeOffset.Now;
            Assert.ThrowsException<NotSupportedException>(() => context.TestDefaultWriteSerializer(cell, dateTime));
        }

        [TestMethod]
        public void TimeSpanOffsetTest()
        {
            var timeSpan = TimeSpan.FromMinutes(288);
            context.TestDefaultWriteSerializer(cell, timeSpan);
            Assert.AreEqual(0.2, cell.Value);
        }

        [TestMethod]
        public void GuidTest()
        {
            var guid = Guid.NewGuid();
            context.TestDefaultWriteSerializer(cell, guid);
            Assert.AreEqual(guid.ToString(), cell.Value);
        }

        [TestMethod]
        public void UriTest()
        {
            var uri = new Uri("http://localhost/home");
            context.TestDefaultWriteSerializer(cell, uri);
            Assert.AreEqual(uri.ToString(), cell.Value);
        }
    }
}
