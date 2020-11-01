using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using Shane32.ExcelLinq.Tests.Models;

namespace Serializers
{
    [TestClass]
    public class DefaultReadSerializer
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

        [TestMethod]
        public void String()
        {
            cell.Value = "test";
            var test = context.TestDefaultReadSerializer(cell, typeof(string));
            Assert.IsInstanceOfType(test, typeof(string));
            Assert.AreEqual("test", test);
        }

        [TestMethod]
        public void StringFromInt()
        {
            cell.Value = 2;
            var test = context.TestDefaultReadSerializer(cell, typeof(string));
            Assert.IsInstanceOfType(test, typeof(string));
            Assert.AreEqual("2", test);
        }

        [TestMethod]
        public void StringFromDouble()
        {
            cell.Value = 2.2;
            var test = context.TestDefaultReadSerializer(cell, typeof(string));
            Assert.IsInstanceOfType(test, typeof(string));
            Assert.AreEqual("2.2", test);
        }

        [TestMethod]
        public void Double()
        {
            cell.Value = 2.2;
            var test = context.TestDefaultReadSerializer(cell, typeof(double));
            Assert.IsInstanceOfType(test, typeof(double));
            Assert.AreEqual(2.2, test);
        }

        [TestMethod]
        public void DoubleFromSingle()
        {
            cell.Value = 2.2f;
            var test = context.TestDefaultReadSerializer(cell, typeof(double));
            Assert.IsInstanceOfType(test, typeof(double));
            Assert.AreEqual((double)2.2f, test);
        }

        [TestMethod]
        public void DoubleFromInt()
        {
            cell.Value = 2;
            var test = context.TestDefaultReadSerializer(cell, typeof(double));
            Assert.IsInstanceOfType(test, typeof(double));
            Assert.AreEqual(2.0, test);
        }

        [TestMethod]
        public void Single()
        {
            cell.Value = 2.2f;
            var test = context.TestDefaultReadSerializer(cell, typeof(float));
            Assert.IsInstanceOfType(test, typeof(float));
            Assert.AreEqual(2.2f, test);
        }

        [TestMethod]
        public void SingleFromDouble()
        {
            cell.Value = 2.2;
            var test = context.TestDefaultReadSerializer(cell, typeof(float));
            Assert.IsInstanceOfType(test, typeof(float));
            Assert.AreEqual(2.2f, test);
        }

        [TestMethod]
        public void SingleFromInt()
        {
            cell.Value = 2;
            var test = context.TestDefaultReadSerializer(cell, typeof(float));
            Assert.IsInstanceOfType(test, typeof(float));
            Assert.AreEqual(2f, test);
        }

        [TestMethod]
        public void Int()
        {
            cell.Value = 2;
            var test = context.TestDefaultReadSerializer(cell, typeof(int));
            Assert.IsInstanceOfType(test, typeof(int));
            Assert.AreEqual(2, test);
        }

        [TestMethod]
        public void IntFromDouble()
        {
            cell.Value = 2.2;
            var test = context.TestDefaultReadSerializer(cell, typeof(int));
            Assert.IsInstanceOfType(test, typeof(int));
            Assert.AreEqual(2, test);
        }

        [TestMethod]
        public void IntFromSingle()
        {
            cell.Value = 2.2f;
            var test = context.TestDefaultReadSerializer(cell, typeof(int));
            Assert.IsInstanceOfType(test, typeof(int));
            Assert.AreEqual(2, test);
        }

        [TestMethod]
        public void IntFromString()
        {
            cell.Value = "2";
            var test = context.TestDefaultReadSerializer(cell, typeof(int));
            Assert.IsInstanceOfType(test, typeof(int));
            Assert.AreEqual(2, test);
        }

        [TestMethod]
        public void DateTimeFromOA()
        {
            var dNow = DateTime.Now;
            var dNowOA = dNow.ToOADate();
            cell.Value = dNowOA;
            var test = context.TestDefaultReadSerializer(cell, typeof(DateTime));
            Assert.IsInstanceOfType(test, typeof(DateTime));
            Assert.AreEqual(DateTime.FromOADate(dNowOA), test);
        }

        [TestMethod]
        public void DateTimeOffsetFromOA()
        {
            cell.Value = DateTime.Now.ToOADate();
            Assert.ThrowsException<NotSupportedException>(() => context.TestDefaultReadSerializer(cell, typeof(DateTimeOffset)));
        }

        [TestMethod]
        public void DateTimeFromString()
        {
            var dNowStr = DateTime.Now.ToShortDateString();
            cell.Value = dNowStr;
            var test = context.TestDefaultReadSerializer(cell, typeof(DateTime));
            Assert.IsInstanceOfType(test, typeof(DateTime));
            Assert.AreEqual(DateTime.Parse(dNowStr), test);
        }

        [TestMethod]
        public void TimespanFromOA()
        {
            var timeSpan = DateTime.Now.TimeOfDay;
            var timeSpanOA = DateTime.FromOADate(0).Add(timeSpan).ToOADate();
            cell.Value = timeSpanOA;
            var test = context.TestDefaultReadSerializer(cell, typeof(TimeSpan));
            Assert.IsInstanceOfType(test, typeof(TimeSpan));
            Assert.AreEqual(DateTime.FromOADate(timeSpanOA).TimeOfDay, test);
        }

        [TestMethod]
        public void TimespanFromString()
        {
            var timeSpanStr = DateTime.Now.TimeOfDay.ToString();
            cell.Value = timeSpanStr;
            var test = context.TestDefaultReadSerializer(cell, typeof(TimeSpan));
            Assert.IsInstanceOfType(test, typeof(TimeSpan));
            Assert.AreEqual(TimeSpan.Parse(timeSpanStr), test);
        }

        [TestMethod]
        public void UriFromString()
        {
            var url = "http://localhost/home";
            cell.Value = url;
            var test = context.TestDefaultReadSerializer(cell, typeof(Uri));
            Assert.IsInstanceOfType(test, typeof(Uri));
            Assert.AreEqual(new Uri(url), test);
        }

        [TestMethod]
        public void GuidFromString()
        {
            var guid = Guid.NewGuid();
            var guidStr = guid.ToString();
            cell.Value = guidStr;
            var test = context.TestDefaultReadSerializer(cell, typeof(Guid));
            Assert.IsInstanceOfType(test, typeof(Guid));
            Assert.AreEqual(guid, test);
        }

        [DataTestMethod]
        [DataRow(true, true, DisplayName = "boolean")]
        [DataRow(0, false, DisplayName = "zero")]
        [DataRow(1, true, DisplayName = "one")]
        [DataRow(-1, true, DisplayName = "minus one")]
        [DataRow(0.0, false, DisplayName = "zero double")]
        [DataRow(1.0, true, DisplayName = "one double")]
        [DataRow(-1.0, true, DisplayName = "minus one double")]
        [DataRow("true", true, DisplayName = "string true")]
        [DataRow("TRUE", true, DisplayName = "string TRUE")]
        [DataRow("True", true, DisplayName = "string True")]
        [DataRow("false", false, DisplayName = "string false")]
        [DataRow("FALSE", false, DisplayName = "string FALSE")]
        [DataRow("False", false, DisplayName = "string False")]
        [DataRow("yes", true, DisplayName = "string yes")]
        [DataRow("YES", true, DisplayName = "string YES")]
        [DataRow("y", true, DisplayName = "string y")]
        [DataRow("Y", true, DisplayName = "string Y")]
        [DataRow("no", false, DisplayName = "string no")]
        [DataRow("NO", false, DisplayName = "string NO")]
        [DataRow("n", false, DisplayName = "string n")]
        [DataRow("N", false, DisplayName = "string N")]
        public void Boolean(object input, bool expected)
        {
            cell.Value = input;
            var test = context.TestDefaultReadSerializer(cell, typeof(bool));
            Assert.IsInstanceOfType(test, typeof(bool));
            Assert.AreEqual(expected, test);
        }
    }
}
