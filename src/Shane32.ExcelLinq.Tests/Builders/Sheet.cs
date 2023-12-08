using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using Shane32.ExcelLinq.Builders;
using Shane32.ExcelLinq.Tests.Models;

namespace Builders
{
    [TestClass]
    public class Sheet
    {
        [TestMethod]
        public void AddSheet()
        {
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class1>();
            var model = builder.Build();
            Assert.AreEqual(1, model.Sheets.Count);
            var sheetModel = model.Sheets[0];
            Assert.AreEqual("Class1", sheetModel.Name);
            Assert.AreEqual(0, sheetModel.AlternateNames.Count);
            Assert.AreEqual(sheetModel, model.Sheets["Class1"]);
            Assert.AreEqual(sheetModel, model.Sheets["CLASS1"]);
            Assert.AreEqual(sheetModel, model.Sheets[typeof(Class1)]);
            Assert.ThrowsException<KeyNotFoundException>(() => model.Sheets["Sheet1"]);
        }

        [TestMethod]
        public void AddSheetAltName()
        {
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class1>("Sheet1");
            var model = builder.Build();
            Assert.AreEqual(1, model.Sheets.Count);
            var sheet = model.Sheets[0];
            Assert.AreEqual("Sheet1", sheet.Name);
            Assert.AreEqual(0, sheet.AlternateNames.Count);
            Assert.AreEqual(sheet, model.Sheets["Sheet1"]);
            Assert.AreEqual(sheet, model.Sheets[typeof(Class1)]);
            Assert.ThrowsException<KeyNotFoundException>(() => model.Sheets["Class1"]);
        }

        [TestMethod]
        public void AddSheetMultipleNames()
        {
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class1>("Sheet1")
                .AlternateName("Sheet2");
            var model = builder.Build();
            Assert.AreEqual(1, model.Sheets.Count);
            var sheet = model.Sheets[0];
            Assert.AreEqual("Sheet1", sheet.Name);
            Assert.AreEqual(1, sheet.AlternateNames.Count);
            Assert.AreEqual("Sheet2", sheet.AlternateNames[0]);
            Assert.AreEqual(sheet, model.Sheets["Sheet1"]);
            Assert.AreEqual(sheet, model.Sheets["SHEET1"]);
            Assert.AreEqual(sheet, model.Sheets["Sheet2"]);
            Assert.AreEqual(sheet, model.Sheets["SHEET2"]);
            Assert.AreEqual(sheet, model.Sheets[typeof(Class1)]);
            Assert.ThrowsException<KeyNotFoundException>(() => model.Sheets["Class1"]);
        }

        [TestMethod]
        public void CantAddNullSheetName()
        {
            var builder = new ExcelModelBuilder();
            Assert.ThrowsException<ArgumentNullException>(() => { builder.Sheet<Class1>(null); });
            var sheet = builder.Sheet<Class1>();
            Assert.ThrowsException<ArgumentNullException>(() => { sheet.AlternateName(null); });
        }

        [TestMethod]
        public void CantAddDuplicateSheetName()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>("Sheet1");
            Assert.ThrowsException<ArgumentException>(() => {
                sheetBuilder.AlternateName("Sheet1");
            });
            Assert.ThrowsException<ArgumentException>(() => {
                sheetBuilder.AlternateName("SHEET1");
            });
        }

        [TestMethod]
        public void CanAddDuplicateSheetClass()
        {
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class1>("Sheet1");
            builder.Sheet<Class1>("Sheet2");
        }

        [TestMethod]
        public void AccessSheetTwice()
        {
            var builder = new ExcelModelBuilder();
            var sheet1 = builder.Sheet<Class1>();
            var sheet2 = builder.Sheet<Class1>();
            Assert.AreEqual(sheet1, sheet2);
            var sheet3 = builder.Sheet<Class2>("Sheet2");
            var sheet4 = builder.Sheet<Class2>("Sheet2");
            Assert.AreEqual(sheet3, sheet4);
        }

        [TestMethod]
        public void AddMultipleSheets()
        {
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class1>();
            builder.Sheet<Class2>();
            var model = builder.Build();
            Assert.AreEqual(2, model.Sheets.Count);
            var sheetModel = model.Sheets[0];
            Assert.AreEqual("Class1", sheetModel.Name);
            Assert.AreEqual(0, sheetModel.AlternateNames.Count);
            Assert.AreEqual(sheetModel, model.Sheets["Class1"]);
            Assert.AreEqual(sheetModel, model.Sheets["CLASS1"]);
            Assert.AreEqual(sheetModel, model.Sheets[typeof(Class1)]);
            sheetModel = model.Sheets[1];
            Assert.AreEqual("Class2", sheetModel.Name);
            Assert.AreEqual(0, sheetModel.AlternateNames.Count);
            Assert.AreEqual(sheetModel, model.Sheets["Class2"]);
            Assert.AreEqual(sheetModel, model.Sheets["CLASS2"]);
            Assert.AreEqual(sheetModel, model.Sheets[typeof(Class2)]);
            Assert.ThrowsException<KeyNotFoundException>(() => model.Sheets["Sheet1"]);
        }

        [TestMethod]
        public void DefaultSheetProperties()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            var model = builder.Build();
            Assert.AreEqual(0, model.Sheets[0].Columns.Count);
            Assert.AreEqual(false, model.Sheets[0].Optional);
            Assert.AreEqual(false, model.Sheets[0].SkipEmptyRows);
            Assert.AreEqual(null, model.Sheets[0].ReadRangeLocator);
            Assert.AreEqual(null, model.Sheets[0].WriteRangeLocator);
            Assert.AreEqual(null, model.Sheets[0].WritePolisher);
        }

        [TestMethod]
        public void SheetProperties()
        {
            Func<ExcelWorksheet, ExcelRange> readRangeLocator = worksheet => worksheet.Cells;
            Func<ExcelWorksheet, ExcelRange> writeRangeLocator = worksheet => worksheet.Cells[1, 1];
            Action<ExcelWorksheet, ExcelRange> writePolisher = (worksheet, range) => { };
            var builder = new ExcelModelBuilder();
            builder.Sheet<Class1>()
                .Optional()
                .SkipEmptyRows()
                .ReadRangeLocator(readRangeLocator)
                .WriteRangeLocator(writeRangeLocator)
                .WritePolisher(writePolisher);
            var model = builder.Build();
            Assert.AreEqual(0, model.Sheets[0].Columns.Count);
            Assert.AreEqual(true, model.Sheets[0].Optional);
            Assert.AreEqual(true, model.Sheets[0].SkipEmptyRows);
            Assert.AreEqual(readRangeLocator, model.Sheets[0].ReadRangeLocator);
            Assert.AreEqual(writeRangeLocator, model.Sheets[0].WriteRangeLocator);
            Assert.AreEqual(writePolisher, model.Sheets[0].WritePolisher);
        }

        [TestMethod]
        public void CtorRedundantChecks()
        {
            Assert.ThrowsException<ArgumentNullException>(() => new SheetModelBuilder<Class1>(null, "test"));
            var builder = new ExcelModelBuilder();
            Assert.ThrowsException<ArgumentNullException>(() => new SheetModelBuilder<Class1>(builder, null));
        }
    }
}
