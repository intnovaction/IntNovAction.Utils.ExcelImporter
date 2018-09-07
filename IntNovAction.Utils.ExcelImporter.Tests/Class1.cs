using System;
using FluentAssertions;
using IntNovAction.Utils.Importer.Tests.SampleClasses;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IntNovAction.Utils.Importer.Tests
{
    [TestClass]
    public class Class1
    {
        [TestMethod]
        public void Test1()
        {
            var importer = new Importer<SampleImportInto>();
        }


        private ReadExcelFile()
        {
            var excelPath = Path.Combine(System.Environment.CurrentDirectory, "SampleExcel.xlsx");
        }
    }
}
