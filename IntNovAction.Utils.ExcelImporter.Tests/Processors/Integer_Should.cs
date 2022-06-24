using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelImporter.CellProcessors;
using IntNovAction.Utils.ExcelImporter;
using IntNovAction.Utils.ExcelImporter.Tests.SampleClasses;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Reflection;

namespace IntNovAction.Utils.ExcelImporter.Tests.Processors
{
    [TestClass]
    public class Integer_Should
    {
        private readonly PropertyInfo IntegerProperty;

        private NumberCellProcessor<int, SampleImportInto> Processor;
        private ImportResult<SampleImportInto> ImportResult;
        private SampleImportInto ObjectToBeFilled;

        public IXLCell Cell { get; private set; }

        public Integer_Should()
        {
            IntegerProperty = typeof(SampleImportInto).GetProperty(nameof(SampleImportInto.IntColumn));
        }

        [TestInitialize()]
        public void Initializer()
        {
            this.Processor = new NumberCellProcessor<int, SampleImportInto>();

            this.ImportResult = new ImportResult<SampleImportInto>();
            this.ObjectToBeFilled = new SampleImportInto();
            this.Cell = GetXLCell();

        }


        [TestMethod]
        public void Process_Integer_Ok()
        {

            Cell.Value = 1;
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, IntegerProperty, Cell);

            cellProcessResult.Should().BeTrue();
            ImportResult.Errors.Should().BeNullOrEmpty();
            ObjectToBeFilled.IntColumn.Should().Be(1);
        }

        [TestMethod]
        public void Process_Letter_AsError()
        {
            Cell.Value = "S";
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, IntegerProperty, Cell);

            cellProcessResult.Should().BeFalse();
            ImportResult.Errors.Should().NotBeNullOrEmpty();
            ImportResult.Errors.Count.Should().Be(1);
            ImportResult.Errors[0].Column.Should().Be(1);
            ImportResult.Errors[0].Row.Should().Be(1);
            ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

            ObjectToBeFilled.IntColumn.Should().Be(0);
        }

        [TestMethod]
        public void Process_EmptyString_AsError()
        {
            Cell.Value = "";
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, IntegerProperty, Cell);

            cellProcessResult.Should().BeFalse();
            ImportResult.Errors.Should().NotBeNullOrEmpty();
            ImportResult.Errors.Count.Should().Be(1);
            ImportResult.Errors[0].Column.Should().Be(1);
            ImportResult.Errors[0].Row.Should().Be(1);
            ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

            ObjectToBeFilled.IntColumn.Should().Be(0);
        }

        [TestMethod]
        public void Process_Null_AsError()
        {
            Cell.Value = null;
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, IntegerProperty, Cell);

            cellProcessResult.Should().BeFalse();
            ImportResult.Errors.Should().NotBeNullOrEmpty();
            ImportResult.Errors.Count.Should().Be(1);
            ImportResult.Errors[0].Column.Should().Be(1);
            ImportResult.Errors[0].Row.Should().Be(1);
            ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

            ObjectToBeFilled.IntColumn.Should().Be(0);
        }

        public IXLCell GetXLCell()
        {
            var wk = new XLWorkbook();
            wk.AddWorksheet("Test1");
            return wk.Worksheet(1).Cell(1, 1);
        }


    }
}
