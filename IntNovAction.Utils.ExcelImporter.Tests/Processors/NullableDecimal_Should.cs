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
    public class NullableDecimal_Should
    {
        private readonly PropertyInfo NullableDecimalProperty;

        private NumberNullableCellProcessor<decimal, SampleImportInto> Processor;
        private ImportResult<SampleImportInto> ImportResult;
        private SampleImportInto ObjectToBeFilled;

        public IXLCell Cell { get; private set; }

        public NullableDecimal_Should()
        {
            NullableDecimalProperty = typeof(SampleImportInto).GetProperty(nameof(SampleImportInto.NullableDecimalColumn));
        }

        [TestInitialize()]
        public void Initializer()
        {
            this.Processor = new NumberNullableCellProcessor<decimal, SampleImportInto>();

            this.ImportResult = new ImportResult<SampleImportInto>();
            this.ObjectToBeFilled = new SampleImportInto();
            this.Cell = GetXLCell();

        }


        [TestMethod]
        public void Process_Decimal_Ok()
        {

            Cell.Value = 1;
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, NullableDecimalProperty, Cell);

            cellProcessResult.Should().BeTrue();
            ImportResult.Errors.Should().BeNullOrEmpty();
            ObjectToBeFilled.NullableDecimalColumn.Should().Be(1);
        }

        [TestMethod]
        public void Process_Letter_AsError()
        {
            Cell.Value = "S";
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, NullableDecimalProperty, Cell);

            cellProcessResult.Should().BeFalse();

            ImportResult.Errors.Should().NotBeNullOrEmpty();
            ImportResult.Errors.Count.Should().Be(1);
            ImportResult.Errors[0].Column.Should().Be(1);
            ImportResult.Errors[0].Row.Should().Be(1);
            ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

            ObjectToBeFilled.NullableDecimalColumn.Should().Be(null);
        }

        [TestMethod]
        public void Process_EmptyString_AsNull()
        {
            Cell.Value = "";
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, NullableDecimalProperty, Cell);

            cellProcessResult.Should().BeTrue();
            ImportResult.Errors.Should().BeEmpty();

            ObjectToBeFilled.NullableDecimalColumn.Should().Be(null);
        }

        [TestMethod]
        public void Process_Null_AsNull()
        {
            Cell.Value = null;
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, NullableDecimalProperty, Cell);

            cellProcessResult.Should().BeTrue();
            ImportResult.Errors.Should().BeEmpty();

            ObjectToBeFilled.NullableDecimalColumn.Should().Be(null);
        }

        public IXLCell GetXLCell()
        {
            var wk = new XLWorkbook();
            wk.AddWorksheet("Test1");
            return wk.Worksheet(1).Cell(1, 1);
        }


    }
}
