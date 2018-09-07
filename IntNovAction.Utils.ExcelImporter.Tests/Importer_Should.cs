using System;
using System.IO;
using System.Linq;
using System.Reflection;
using FluentAssertions;
using IntNovAction.Utils.Importer.Tests.SampleClasses;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IntNovAction.Utils.Importer.Tests
{
    [TestClass]
    public class Importer_Should
    {
        [TestMethod]
        public void Import_Int_Column_From_Name()
        {
            var importer = new Importer<SampleImportInto>();

            using (var stream = OpenExcel())
            {
                var lista = importer
                    .FromExcel(stream)
                    .For(p => p.IntColumn, "Int Column")
                    .Import();

                lista.Result.Should().Be(ImportErrorResult.Ok);

                lista.Errors.Should().NotBeNull();
                lista.Errors.Should().BeEmpty();

                lista.ImportedItems.Should().NotBeNull();
                lista.ImportedItems.Count().Should().Be(5);
            }
        }

        [TestMethod]
        public void Import_NullableInt_Column_From_Name()
        {
            var importer = new Importer<SampleImportInto>();

            using (var stream = OpenExcel())
            {
                var lista = importer
                    .FromExcel(stream)
                    .For(p => p.NullableIntColumn, "Nullable Int Column")
                    .Import();

                lista.Result.Should().Be(ImportErrorResult.Ok);

                lista.Errors.Should().NotBeNull();
                lista.Errors.Should().BeEmpty();

                lista.ImportedItems.Should().NotBeNull();
                lista.ImportedItems.Count().Should().Be(5);

                lista.ImportedItems[0].NullableIntColumn.Should().Be(1);
                lista.ImportedItems[1].NullableIntColumn.Should().BeNull();
            }
        }

        public Stream OpenExcel()
        {
            var stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("IntNovAction.Utils.ExcelImporter.Tests.SampleExcels.SampleExcel.xlsx");

            return stream;
        }

    }
}
