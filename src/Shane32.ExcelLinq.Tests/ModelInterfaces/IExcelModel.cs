using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shane32.ExcelLinq;
using Shane32.ExcelLinq.Builders;
using Shane32.ExcelLinq.Models;
using Shane32.ExcelLinq.Tests.Models;

namespace ModelInterfaces
{
    [TestClass]
    public class IExcelModel
    {
        [TestMethod]
        public void InvalidSheetType()
        {
            var context = new MyContext();
            Assert.ThrowsException<KeyNotFoundException>(() => context.Model.Sheets[typeof(Class3)]);
            var found = context.Model.Sheets.TryGetValue(typeof(Class3), out var value);
            Assert.IsFalse(found);
        }

        [TestMethod]
        public void InvalidSheetName()
        {
            var context = new MyContext();
            Assert.ThrowsException<KeyNotFoundException>(() => context.Model.Sheets["Class3"]);
            var found = context.Model.Sheets.TryGetValue("Class3", out var value);
            Assert.IsFalse(found);
        }

        [TestMethod]
        public void NullSheetType()
        {
            var context = new MyContext();
            Assert.ThrowsException<ArgumentNullException>(() => context.Model.Sheets[(Type)null]);
            Assert.ThrowsException<ArgumentNullException>(() => {
                context.Model.Sheets.TryGetValue((Type)null, out var value);
            });
        }

        [TestMethod]
        public void NullSheetName()
        {
            var context = new MyContext();
            Assert.ThrowsException<ArgumentNullException>(() => context.Model.Sheets[(string)null]);
            Assert.ThrowsException<ArgumentNullException>(() => {
                context.Model.Sheets.TryGetValue((string)null, out var value);
            });
        }

        [TestMethod]
        public void EnumerateSheets()
        {
            var context = new MyContext();
            var model = context.Model;
            var sheets = model.Sheets.ToList();
            var enumerator = ((IEnumerable)model.Sheets).GetEnumerator();
            var testCount = 0;
            while (enumerator.MoveNext()) testCount++;
            Assert.AreEqual(2, testCount);
            Assert.AreEqual(2, sheets.Count);
            Assert.IsNotNull(sheets[0]);
            Assert.AreEqual("Class1", sheets[0].Name);
            Assert.AreEqual("AltName1", sheets[0].AlternateNames.Single());
            Assert.IsNotNull(sheets[1]);
            Assert.AreEqual("Sheet2", sheets[1].Name);
        }

        [TestMethod]
        public void FindSheetByType()
        {
            var context = new MyContext();
            var model = context.Model;
            var sheets = model.Sheets;
            var sheet1 = sheets[typeof(Class1)];
            Assert.IsNotNull(sheet1);
            Assert.AreEqual("Class1", sheet1.Name);
            var sheet2 = sheets[typeof(Class2)];
            Assert.IsNotNull(sheet2);
            Assert.AreEqual("Sheet2", sheet2.Name);
            Assert.IsTrue(sheets.TryGetValue(typeof(Class1), out var sheet1b));
            Assert.AreEqual(sheet1, sheet1b);
            Assert.IsTrue(sheets.TryGetValue(typeof(Class2), out var sheet2b));
            Assert.AreEqual(sheet2, sheet2b);
        }

        [TestMethod]
        public void FindSheetByName()
        {
            var context = new MyContext();
            var model = context.Model;
            var sheets = model.Sheets;
            var sheet1 = sheets["Class1"];
            Assert.IsNotNull(sheet1);
            Assert.AreEqual("Class1", sheet1.Name);
            var sheet1c = sheets["AltName1"];
            Assert.IsNotNull(sheet1c);
            Assert.AreEqual(sheet1, sheet1c);
            var sheet2 = sheets["Sheet2"];
            Assert.IsNotNull(sheet2);
            Assert.AreEqual("Sheet2", sheet2.Name);
            Assert.IsTrue(sheets.TryGetValue("Class1", out var sheet1b));
            Assert.AreEqual(sheet1, sheet1b);
            Assert.IsTrue(sheets.TryGetValue("AltName1", out var sheet1d));
            Assert.AreEqual(sheet1, sheet1d);
            Assert.IsTrue(sheets.TryGetValue("Sheet2", out var sheet2b));
            Assert.AreEqual(sheet2, sheet2b);
        }

        private class MyContext : ExcelContext
        {
            protected override void OnModelCreating(ExcelModelBuilder modelBuilder)
            {
                modelBuilder.Sheet<Class1>()
                    .AlternateName("AltName1");
                modelBuilder.Sheet<Class2>("Sheet2");
            }
        }

    }
}
