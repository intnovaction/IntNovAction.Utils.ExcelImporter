using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelImporter.CellProcessors;
using IntNovAction.Utils.Importer;
using IntNovAction.Utils.Importer.Tests.SampleClasses;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Reflection;

namespace IntNovAction.Utils.ExcelImporter.Tests.Processors
{
    [TestClass]
    public class String_Should
    {
        private readonly PropertyInfo StringProperty;

        private StringCellProcessor<SampleImportInto> Processor;
        private ImportResult<SampleImportInto> ImportResult;
        private SampleImportInto ObjectToBeFilled;

        public IXLCell Cell { get; private set; }

        public String_Should()
        {
            StringProperty = typeof(SampleImportInto).GetProperty(nameof(SampleImportInto.StringColumn));
        }

        [TestInitialize()]
        public void Initializer()
        {
            this.Processor = new StringCellProcessor<SampleImportInto>();

            this.ImportResult = new ImportResult<SampleImportInto>();
            this.ObjectToBeFilled = new SampleImportInto();
            this.Cell = GetXLCell();

        }


        [TestMethod]
        public void Process_String_Ok()
        {

            Cell.Value = "Hello";
            var cellProcessResult = this.Processor.SetValue(ImportResult, ObjectToBeFilled, StringProperty, Cell);

            cellProcessResult.Should().BeTrue();
            ImportResult.Errors.Should().BeNullOrEmpty();
            ObjectToBeFilled.StringColumn.Should().Be("Hello");
        }


        [TestMethod]
        public void Process_Null_AsNull()
        {
            Cell.Value = null;
            var cellProcessResult = this.Processor.SetValue(ImportResult, ObjectToBeFilled, StringProperty, Cell);

            cellProcessResult.Should().BeTrue();
            ImportResult.Errors.Should().BeEmpty();

            ObjectToBeFilled.StringColumn.Should().Be(null);
        }

        public IXLCell GetXLCell()
        {
            var wk = new XLWorkbook();
            wk.AddWorksheet("Test1");
            return wk.Worksheet(1).Cell(1, 1);
        }


    }
}
