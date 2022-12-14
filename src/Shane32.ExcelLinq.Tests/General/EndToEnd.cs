using System;
using System.IO;
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
    public class EndToEnd
    {
        [TestMethod]
        public void NullConstructorsThrow()
        {
            Assert.ThrowsException<ArgumentNullException>(() => new TestFileContext((string)null));
            Assert.ThrowsException<ArgumentNullException>(() => new TestFileContext((Stream)null));
            Assert.ThrowsException<ArgumentNullException>(() => new TestFileContext((ExcelPackage)null));
        }


        [TestMethod]
        public void ReadSampleCsvFile()
        {
            var xl = new TestFileContext();
            using var stream = new System.IO.FileStream("test.csv", System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read);
            xl.ReadCsv(stream);
            ReadSample1File_test(xl);
        }

        [TestMethod]
        public void ReadSample1File()
        {
            using var stream = new System.IO.FileStream("test.xlsx", System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read);
            var xl = new TestFileContext(stream);
            ReadSample1File_test(xl);
        }

        [TestMethod]
        public void ReadSample1File_2()
        {
            var xl = new TestFileContext("test.xlsx");
            ReadSample1File_test(xl);
        }

        [TestMethod]
        public void ReadSample1File_3()
        {
            var xl = new TestFileContext(new ExcelPackage(new FileInfo("test.xlsx")));
            ReadSample1File_test(xl);
        }

        [TestMethod]
        public void ReadAndWrite()
        {
            var xl = new TestFileContext("test.xlsx");
            xl.SerializeToFile("test-out.xlsx");
            xl.SerializeToStream();
            xl.SerializeToExcelPackage();
            var xl2 = new TestFileContext("test-out.xlsx");
            Assert.AreEqual(xl.GetSheet<TestFileContext.Sheet1>().Count, xl2.GetSheet<TestFileContext.Sheet1>().Count);
            Assert.AreEqual(xl.GetSheet<Class1>().Count, xl2.GetSheet<Class1>().Count);
        }

        public void ReadSample1File_test(TestFileContext xl) {
            var sheet1 = xl.GetSheet<TestFileContext.Sheet1>();
            var sheet2 = xl.GetSheet<Class1>();
            //Assert.AreEqual(2, sheet1.Count);
            var s1row = sheet1[0];

            Assert.AreEqual(DateTime.Parse("7/1/2020"), s1row.Date);
            Assert.AreEqual(52, s1row.Quantity);
            Assert.AreEqual("Widgets", s1row.Description);
            Assert.AreEqual(45.99m, s1row.Amount);
            Assert.AreEqual(2391.48m, s1row.Total);
            Assert.AreEqual(null, s1row.Notes);

            s1row = sheet1[1];
            Assert.AreEqual(DateTime.Parse("7/23/2020"), s1row.Date);
            Assert.AreEqual(22, s1row.Quantity);
            Assert.AreEqual("Bolts", s1row.Description);
            Assert.AreEqual(2.54m, s1row.Amount);
            Assert.AreEqual(55.88m, s1row.Total);
            Assert.AreEqual("Each bolt is a set of two", s1row.Notes);

            Assert.AreEqual(9, sheet2.Count);

            var s2row = sheet2[0];
            Assert.AreEqual(1, s2row.IntColumn);
            Assert.AreEqual(1f, s2row.FloatColumn);
            Assert.AreEqual(1.0, s2row.DoubleColumn);
            Assert.AreEqual("test", s2row.StringColumn);
            Assert.AreEqual(true, s2row.BooleanColumn);
            Assert.AreEqual(DateTime.Parse("8/2/2020"), s2row.DateTimeColumn);
            Assert.AreEqual(TimeSpan.FromHours(14), s2row.TimeSpanColumn);
            Assert.AreEqual(new Uri("http://localhost/test"), s2row.UriColumn);
            Assert.AreEqual(Guid.Parse("f1dc7e7d-d63e-4279-8dfd-cecb6e26cda8"), s2row.GuidColumn);
            Assert.AreEqual(3, s2row.NullableIntColumn);

            s2row = sheet2[1];
            Assert.AreEqual(1, s2row.IntColumn);
            Assert.AreEqual(1.1f, s2row.FloatColumn);
            Assert.AreEqual(1.1, s2row.DoubleColumn);
            Assert.AreEqual("test2", s2row.StringColumn);
            Assert.AreEqual(false, s2row.BooleanColumn);
            Assert.AreEqual(DateTime.Parse("8/1/2020"), s2row.DateTimeColumn);
            Assert.AreEqual(TimeSpan.FromHours(14), s2row.TimeSpanColumn);
            Assert.AreEqual(new Uri("http://localhost/help"), s2row.UriColumn);
            Assert.AreEqual(Guid.Parse("89892480-4179-42c7-9e2f-e0bb1094dd6b"), s2row.GuidColumn);
            Assert.AreEqual(3, s2row.NullableIntColumn);

            s2row = sheet2[2];
            Assert.AreEqual(1, s2row.IntColumn);
            Assert.AreEqual(1.1f, s2row.FloatColumn);
            Assert.AreEqual(1.1, s2row.DoubleColumn);
            Assert.AreEqual("test2", s2row.StringColumn);
            Assert.AreEqual(true, s2row.BooleanColumn);
            Assert.AreEqual(DateTime.Parse("8/3/2020"), s2row.DateTimeColumn);
            Assert.AreEqual(TimeSpan.FromHours(14), s2row.TimeSpanColumn);
            Assert.AreEqual(new Uri("http://localhost/help"), s2row.UriColumn);
            Assert.AreEqual(Guid.Parse("89892480-4179-42c7-9e2f-e0bb1094dd6b"), s2row.GuidColumn);
            Assert.AreEqual(3, s2row.NullableIntColumn);

            s2row = sheet2[3];
            Assert.AreEqual(1, s2row.IntColumn);
            Assert.AreEqual(1.1f, s2row.FloatColumn);
            Assert.AreEqual(1.1, s2row.DoubleColumn);
            Assert.AreEqual("test2", s2row.StringColumn);
            Assert.AreEqual(true, s2row.BooleanColumn);
            Assert.AreEqual(DateTime.Parse("8/1/2020 2:30 PM"), s2row.DateTimeColumn);
            Assert.AreEqual(TimeSpan.Parse("2:34:56"), s2row.TimeSpanColumn);
            Assert.AreEqual(new Uri("http://localhost/help"), s2row.UriColumn);
            Assert.AreEqual(Guid.Parse("89892480-4179-42c7-9e2f-e0bb1094dd6b"), s2row.GuidColumn);
            Assert.AreEqual(null, s2row.NullableIntColumn);

            Assert.AreEqual(false, sheet2[4].BooleanColumn);
            Assert.AreEqual(true, sheet2[5].BooleanColumn);
            Assert.AreEqual(false, sheet2[6].BooleanColumn);
            Assert.AreEqual(true, sheet2[7].BooleanColumn);
            Assert.AreEqual(false, sheet2[8].BooleanColumn);
        }

        [TestMethod]
        public void ReadDynamicallyGeneratedFile()
        {
            var package = new ExcelPackage();
            var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
            sheet1.SetValue(1, 1, "StringColumn");
            sheet1.SetValue(1, 2, "IntColumn");
            sheet1.SetValue(1, 3, "BoolColumn");
            sheet1.SetValue(2, 1, "test");
            sheet1.SetValue(2, 2, 2.0);
            sheet1.SetValue(2, 3, "true");
            sheet1.SetValue(4, 1, "test2");
            sheet1.SetValue(4, 3, "y");
            var stream = new System.IO.MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            var mock = new Mock<ExcelContext>(stream);
            mock.CallBase = true;
            mock.Protected().Setup("OnModelCreating", ItExpr.IsAny<ExcelModelBuilder>()).Callback<ExcelModelBuilder>(
                (builder) => {
                    var sheetBuilder = builder.Sheet<Class1>("Sheet1");
                    sheetBuilder.Column(x => x.StringColumn);
                    sheetBuilder.Column(x => x.IntColumn).Optional();
                    sheetBuilder.Column(x => x.BooleanColumn, "BoolColumn");
                    sheetBuilder.SkipEmptyRows();
                }).Verifiable();
            var context = mock.Object;
            var sheet1Data = context.GetSheet<Class1>();
            Assert.AreEqual(2, sheet1Data.Count);
            Assert.AreEqual("test", sheet1Data[0].StringColumn);
            Assert.AreEqual(2, sheet1Data[0].IntColumn);
            Assert.AreEqual(true, sheet1Data[0].BooleanColumn);
            Assert.AreEqual("test2", sheet1Data[1].StringColumn);
            Assert.AreEqual(true, sheet1Data[1].BooleanColumn);
            mock.Verify();
        }
    }
}
