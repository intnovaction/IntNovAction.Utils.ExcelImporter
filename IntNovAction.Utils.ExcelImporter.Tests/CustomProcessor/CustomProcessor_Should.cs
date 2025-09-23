using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelImporter.Tests.SampleClasses;
using IntNovAction.Utils.ExcelImporter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System;

namespace IntNovAction.Utils.ExcelImporter.Tests.CustomProcessor
{
    [TestClass]
    public class CustomProcessor_Should
    {
        

        public Stream OpenExcel()
        {
            var stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("IntNovAction.Utils.ExcelImporter.Tests.SampleExcels.SampleExcel.xlsx");

            return stream;
        }

        public Stream OpenExcelDupColumns()
        {
            var stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("IntNovAction.Utils.ExcelImporter.Tests.SampleExcels.SampleExcel_DupColumn.xlsx");

            return stream;
        }




        [TestMethod]
        public void Execute_Custom_For_Should_Invoke_Custom_Action()
        {
            // Arrange
            var importer = new Importer<SampleImportInto>();
            bool customForCalled = false;
            

            using (var stream = OpenExcel())
            {
                var lista = importer
                    .SetDuplicatedColumnsStrategy(DuplicatedColumnStrategy.RaiseError)
                    .FromExcel(stream)
                    .SetErrorStrategy(ErrorStrategy.AddElement)
                    .CustomFor((Dictionary<string, string> rowValues, SampleImportInto destination) =>
                    {
                        customForCalled = true;
                    })
                    .Import();

                // Assert
                customForCalled.Should().BeTrue("CustomFor action should be called during import");
            }
        }

        [TestMethod]
        public void Take_First_Value_When_Columns_Are_Duplicated_And_Strategy_Set()
        {
            var importer = new Importer<SampleImportInto>();

            var intColumReadVal = 0;
            using (var stream = OpenExcelDupColumns())
            {
                var lista = importer
                    .SetDuplicatedColumnsStrategy(DuplicatedColumnStrategy.TakeFirst)
                    .FromExcel(stream, "Duplicated Columns")
                    .SetErrorStrategy(ErrorStrategy.AddElement)
                    .For(p => p.IntColumn, "Int Column")
                    .CustomFor((Dictionary<string, string> rowValues, SampleImportInto destination) =>
                    {
                        intColumReadVal = Int32.Parse(rowValues["Int Column"]);
                    })
                    .Import();

                lista.Result.Should().Be(ImportErrorResult.PartialOk);

                lista.Errors.Should().NotBeNullOrEmpty();

                lista.ImportedItems.Should().NotBeNullOrEmpty();
                intColumReadVal.Should().Be(lista.ImportedItems.Last().IntColumn);
            }
        }

        [TestMethod]
        public void Take_Last_Value_When_Columns_Are_Duplicated_And_Strategy_Set()
        {
            var importer = new Importer<SampleImportInto>();
            var intColumReadVal = 0;

            using (var stream = OpenExcelDupColumns())
            {
                var lista = importer
                    .SetDuplicatedColumnsStrategy(DuplicatedColumnStrategy.TakeLast)
                    .FromExcel(stream, "Duplicated Columns")
                    .SetErrorStrategy(ErrorStrategy.AddElement)
                    .For(p => p.IntColumn, "Int Column")
                    .CustomFor((Dictionary<string, string> rowValues, SampleImportInto destination) =>
                    {
                        intColumReadVal = Int32.Parse(rowValues["Int Column"]);
                    })
                    .Import();

                lista.Result.Should().Be(ImportErrorResult.PartialOk);

                lista.Errors.Should().NotBeNullOrEmpty();

                lista.ImportedItems.Should().NotBeNullOrEmpty();
                
                intColumReadVal.Should().Be(lista.ImportedItems.Last().IntColumn);
            }
        }
    }
}
