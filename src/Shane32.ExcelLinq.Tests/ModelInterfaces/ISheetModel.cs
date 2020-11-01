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
    public class ISheetModel
    {
        [TestMethod]
        public void EnumerateColumns()
        {
            var sheetModel = new MyContext().Model.Sheets[typeof(Class1)];
            var enumerator = ((IEnumerable)sheetModel.Columns).GetEnumerator();
            var testCount = 0;
            while (enumerator.MoveNext()) testCount++;
            Assert.AreEqual(2, testCount);
            var columns = sheetModel.Columns.ToList();
            Assert.AreEqual(2, columns.Count);
            Assert.IsNotNull(columns[0]);
            Assert.AreEqual("StringColumn", columns[0].Name);
            Assert.AreEqual("String", columns[0].AlternateNames.Single());
            Assert.IsNotNull(columns[1]);
            Assert.AreEqual("IntColumn", columns[1].Name);
        }

        [TestMethod]
        public void FindColumnByName()
        {
            var context = new MyContext();
            var columns = context.Model.Sheets.Single().Columns;
            var column1 = columns["StringColumn"];
            Assert.IsNotNull(column1);
            Assert.AreEqual("StringColumn", column1.Name);
            var column1c = columns["String"];
            Assert.IsNotNull(column1c);
            Assert.AreEqual(column1, column1c);
            var column2 = columns["IntColumn"];
            Assert.IsNotNull(column2);
            Assert.AreEqual("IntColumn", column2.Name);
            Assert.IsTrue(columns.TryGetValue("StringColumn", out var column1b));
            Assert.AreEqual(column1, column1b);
            Assert.IsTrue(columns.TryGetValue("String", out var column1d));
            Assert.AreEqual(column1, column1d);
            Assert.IsTrue(columns.TryGetValue("IntColumn", out var column2b));
            Assert.AreEqual(column2, column2b);
        }

        [TestMethod]
        public void InvalidColumnName()
        {
            var context = new MyContext();
            var columns = context.Model.Sheets["Class1"].Columns;
            Assert.ThrowsException<KeyNotFoundException>(() => columns["Invalid"]);
            var found = columns.TryGetValue("Invalid", out var value);
            Assert.IsFalse(found);
        }

        [TestMethod]
        public void NullColumnName()
        {
            var context = new MyContext();
            var columns = context.Model.Sheets["Class1"].Columns;
            Assert.ThrowsException<ArgumentNullException>(() => columns[(string)null]);
            Assert.ThrowsException<ArgumentNullException>(() => {
                columns.TryGetValue((string)null, out var value);
            });
        }

        private class MyContext : ExcelContext
        {
            protected override void OnModelCreating(ExcelModelBuilder modelBuilder)
            {
                var sheetBuilder = modelBuilder.Sheet<Class1>();
                sheetBuilder.Column(x => x.StringColumn)
                    .AlternateName("String");
                sheetBuilder.Column(x => x.IntColumn);
            }
        }

    }
}
