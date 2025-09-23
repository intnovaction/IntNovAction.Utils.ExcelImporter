using System.Reflection;
using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelImporter.CellProcessors;
using IntNovAction.Utils.ExcelImporter;
using IntNovAction.Utils.ExcelImporter.Tests.SampleClasses;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IntNovAction.Utils.ExcelImporter.Tests.Processors
{
    [TestClass]
    public class Decimal_Should
    {
        private readonly PropertyInfo DecimalProperty;

        private NumberCellProcessor<decimal, SampleImportInto> Processor;
        private ImportResult<SampleImportInto> ImportResult;
        private SampleImportInto ObjectToBeFilled;

        public IXLCell Cell { get; private set; }

        public Decimal_Should()
        {
            this.DecimalProperty = typeof(SampleImportInto).GetProperty(nameof(SampleImportInto.DecimalColumn));
        }

        [TestInitialize()]
        public void Initializer()
        {
            this.Processor = new NumberCellProcessor<decimal, SampleImportInto>();

            this.ImportResult = new ImportResult<SampleImportInto>();
            this.ObjectToBeFilled = new SampleImportInto();
            this.Cell = GetXLCell();

        }


        [TestMethod]
        public void Process_Decimal_Ok()
        {

            this.Cell.Value = 1;
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(this.ImportResult, this.ObjectToBeFilled, this.DecimalProperty, this.Cell);

            cellProcessResult.Should().BeTrue();
            this.ImportResult.Errors.Should().BeNullOrEmpty();
            this.ObjectToBeFilled.DecimalColumn.Should().Be(1);
        }

        [TestMethod]
        public void Process_Letter_AsError()
        {
            this.Cell.Value = "S";
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(this.ImportResult, this.ObjectToBeFilled, this.DecimalProperty, this.Cell);

            cellProcessResult.Should().BeFalse();
            this.ImportResult.Errors.Should().NotBeNullOrEmpty();
            this.ImportResult.Errors.Count.Should().Be(1);
            this.ImportResult.Errors[0].Column.Should().Be(1);
            this.ImportResult.Errors[0].Row.Should().Be(1);
            this.ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

            this.ObjectToBeFilled.DecimalColumn.Should().Be(0);
        }

        [TestMethod]
        public void Process_EmptyString_AsError()
        {
            this.Cell.Value = "";
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(this.ImportResult, this.ObjectToBeFilled, this.DecimalProperty, this.Cell);

            cellProcessResult.Should().BeFalse();
            this.ImportResult.Errors.Should().NotBeNullOrEmpty();
            this.ImportResult.Errors.Count.Should().Be(1);
            this.ImportResult.Errors[0].Column.Should().Be(1);
            this.ImportResult.Errors[0].Row.Should().Be(1);
            this.ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

            this.ObjectToBeFilled.DecimalColumn.Should().Be(0);
        }

        [TestMethod]
        public void Process_Null_AsError()
        {
            this.Cell.Value = null;
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(this.ImportResult, this.ObjectToBeFilled, this.DecimalProperty, this.Cell);

            cellProcessResult.Should().BeFalse();
            this.ImportResult.Errors.Should().NotBeNullOrEmpty();
            this.ImportResult.Errors.Count.Should().Be(1);
            this.ImportResult.Errors[0].Column.Should().Be(1);
            this.ImportResult.Errors[0].Row.Should().Be(1);
            this.ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

            this.ObjectToBeFilled.DecimalColumn.Should().Be(0);
        }

        public IXLCell GetXLCell()
        {
            var wk = new XLWorkbook();
            wk.AddWorksheet("Test1");
            return wk.Worksheet(1).Cell(1, 1);
        }


    }
}
