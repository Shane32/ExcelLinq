using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shane32.ExcelLinq;
using Shane32.ExcelLinq.Builders;
using Shane32.ExcelLinq.Tests.Models;

namespace Builders
{
    [TestClass]
    public class Excel
    {
        [TestMethod]
        public void DefaultProperties()
        {
            var builder = new ExcelModelBuilder();
            var model = builder.Build();
            Assert.AreEqual(0, model.Sheets.Count);
            Assert.AreEqual(false, model.IgnoreSheetNames);
        }

        [TestMethod]
        public void SetProperties()
        {
            var builder = new ExcelModelBuilder();
            builder.IgnoreSheetNames();
            var model = builder.Build();
            Assert.AreEqual(true, model.IgnoreSheetNames);
        }
    }
}
