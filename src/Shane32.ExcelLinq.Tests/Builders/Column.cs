using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using Shane32.ExcelLinq.Builders;
using Shane32.ExcelLinq.Tests.Models;

namespace Builders
{
    [TestClass]
    public class Column
    {
        [TestMethod]
        public void AddColumn()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn);
            var model = builder.Build().Sheets[0];
            Assert.AreEqual(1, model.Columns.Count);
            var columnModel = model.Columns[0];
            Assert.AreEqual("StringColumn", columnModel.Name);
            Assert.AreEqual(0, columnModel.AlternateNames.Count);
            Assert.AreEqual(columnModel, model.Columns["StringColumn"]);
            Assert.AreEqual(columnModel, model.Columns["STRINGCOLUMN"]);
            Assert.ThrowsException<KeyNotFoundException>(() => model.Columns["Column1"]);
        }

        [TestMethod]
        public void AddColumnAltName()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn, "Column1");
            var model = builder.Build().Sheets[0];
            Assert.AreEqual(1, model.Columns.Count);
            var columnModel = model.Columns[0];
            Assert.AreEqual("Column1", columnModel.Name);
            Assert.AreEqual(0, columnModel.AlternateNames.Count);
            Assert.AreEqual(columnModel, model.Columns["Column1"]);
            Assert.AreEqual(columnModel, model.Columns["COLUMN1"]);
            Assert.ThrowsException<KeyNotFoundException>(() => model.Columns["StringColumn"]);
        }

        [TestMethod]
        public void AddSheetMultipleNames()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn, "Column1")
                .AlternateName("Column2");
            var model = builder.Build().Sheets[0];
            Assert.AreEqual(1, model.Columns.Count);
            var columnModel = model.Columns[0];
            Assert.AreEqual("Column1", columnModel.Name);
            Assert.AreEqual(1, columnModel.AlternateNames.Count);
            Assert.AreEqual("Column2", columnModel.AlternateNames[0]);
            Assert.AreEqual(columnModel, model.Columns["Column1"]);
            Assert.AreEqual(columnModel, model.Columns["COLUMN1"]);
            Assert.AreEqual(columnModel, model.Columns["Column2"]);
            Assert.AreEqual(columnModel, model.Columns["COLUMN2"]);
            Assert.ThrowsException<KeyNotFoundException>(() => model.Columns["StringColumn"]);
        }

        [TestMethod]
        public void CantAddDuplicateColumnName()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            var columnBuilder = sheetBuilder.Column(x => x.StringColumn, "Column1");
            Assert.ThrowsException<ArgumentException>(() => {
                columnBuilder.AlternateName("Column1");
            });
            Assert.ThrowsException<ArgumentException>(() => {
                columnBuilder.AlternateName("COLUMN1");
            });
        }

        [TestMethod]
        public void CantAddDuplicateColumnTwice()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            var columnBuilder = sheetBuilder.Column(x => x.StringColumn, "Column1");
            Assert.ThrowsException<InvalidOperationException>(() => {
                sheetBuilder.Column(x => x.StringColumn, "Column2");
            });
        }

        [TestMethod]
        public void CantAddNullColumnName()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            Assert.ThrowsException<ArgumentNullException>(() => sheetBuilder.Column(x => x.StringColumn, null));
            var columnBuilder = sheetBuilder.Column(x => x.StringColumn);
            Assert.ThrowsException<ArgumentNullException>(() => columnBuilder.AlternateName(null));
        }

        [TestMethod]
        public void AddMultipleColumns()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn);
            sheetBuilder.Column(x => x.IntColumn);
            var model = builder.Build().Sheets[0];
            Assert.AreEqual(2, model.Columns.Count);

            var columnModel = model.Columns[0];
            Assert.AreEqual("StringColumn", columnModel.Name);
            Assert.AreEqual(0, columnModel.AlternateNames.Count);
            Assert.AreEqual(columnModel, model.Columns["StringColumn"]);
            Assert.AreEqual(columnModel, model.Columns["STRINGCOLUMN"]);

            var columnModel2 = model.Columns[1];
            Assert.AreEqual("IntColumn", columnModel2.Name);
            Assert.AreEqual(0, columnModel2.AlternateNames.Count);
            Assert.AreEqual(columnModel2, model.Columns["IntColumn"]);
            Assert.AreEqual(columnModel2, model.Columns["INTCOLUMN"]);
        }

        [TestMethod]
        public void DefaultColumnProperties()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            var columnBuilder = sheetBuilder.Column(x => x.StringColumn);
            var model = builder.Build();
            var columnModel = model.Sheets[0].Columns[0];
            Assert.AreEqual(false, columnModel.Optional);
            Assert.AreEqual(null, columnModel.ReadSerializer);
            Assert.AreEqual(null, columnModel.WriteSerializer);
            Assert.AreEqual(null, columnModel.WritePolisher);
            Assert.AreEqual(null, columnModel.HeaderFormatter);
            Assert.AreEqual(null, columnModel.ColumnFormatter);
        }

        [TestMethod]
        public void SheetProperties()
        {
            Func<ExcelRange, string> readSerializer = range => "testdata";
            Action<ExcelRange, string> writeSerializer = (range, value) => range.Value = "testwrite";
            Action<ExcelRange> headerFormatter = _ => { };
            Action<ExcelRange> columnFormatter = _ => { };
            Action<ExcelRange> writePolisher = _ => { };

            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class1>();
            sheetBuilder.Column(x => x.StringColumn)
                .Optional()
                .ReadSerializer(readSerializer)
                .WriteSerializer(writeSerializer)
                .HeaderFormatter(headerFormatter)
                .ColumnFormatter(columnFormatter)
                .WritePolisher(writePolisher);
            var model = builder.Build();
            var sheetModel = model.Sheets[0];
            var columnModel = sheetModel.Columns[0];
            Assert.AreEqual(true, columnModel.Optional);
            Assert.AreEqual("testdata", columnModel.ReadSerializer(null));
            var newFile = new ExcelPackage();
            var newSheet = newFile.Workbook.Worksheets.Add("Sheet1");
            var cell = newSheet.Cells[1, 1];
            columnModel.WriteSerializer(cell, "test");
            Assert.AreEqual("testwrite", cell.Value);
            Assert.AreEqual(headerFormatter, columnModel.HeaderFormatter);
            Assert.AreEqual(columnFormatter, columnModel.ColumnFormatter);
            Assert.AreEqual(writePolisher, columnModel.WritePolisher);
        }

        [TestMethod]
        public void ColumnMemberTypes()
        {
            var builder = new ExcelModelBuilder();
            var sheetBuilder = builder.Sheet<Class3>();
            sheetBuilder.Column(x => x.Valid1);
            sheetBuilder.Column(x => x.Valid2);
            sheetBuilder.Column(x => x.Valid3);
            sheetBuilder.Column(x => x.Valid4);
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => sheetBuilder.Column(x => x.Invalid1));
        }
    }
}
