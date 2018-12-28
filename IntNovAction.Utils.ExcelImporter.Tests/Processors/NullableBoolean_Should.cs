using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelImporter.CellProcessors;
using IntNovAction.Utils.Importer;
using IntNovAction.Utils.Importer.Tests.SampleClasses;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Reflection;

namespace IntNovAction.Utils.ExcelImporter.Tests.Processors
{
    [TestClass]
    public class NullableBoolean_Should
    {
        private readonly PropertyInfo BooleanProperty;

        private BooleanCellProcessor<SampleImportInto> Processor;
        private ImportResult<SampleImportInto> ImportResult;
        private SampleImportInto ObjectToBeFilled;

        public IXLCell Cell { get; private set; }

        public NullableBoolean_Should()
        {
            BooleanProperty = typeof(SampleImportInto).GetProperty(nameof(SampleImportInto.BooleanColumn));
        }

        [TestInitialize()]
        public void Initializer()
        {
            var options = new BooleanOptions();

            this.Processor = new BooleanCellProcessor<SampleImportInto>(true, options);

            this.ImportResult = new ImportResult<SampleImportInto>();
            this.ObjectToBeFilled = new SampleImportInto();
            this.Cell = GetXLCell();

        }


        [TestMethod]
        public void Process_Boolean_True_Value()
        {

            Cell.Value = "yes";
            var cellProcessResult = this.Processor.SetValue(ImportResult, ObjectToBeFilled, BooleanProperty, Cell);

            cellProcessResult.Should().BeTrue();

            ImportResult.Errors.Should().BeNullOrEmpty();
            ObjectToBeFilled.BooleanColumn.Should().Be(true);
        }

        [TestMethod]
        public void Process_Boolean_False_Value()
        {

            Cell.Value = "no";
            var cellProcessResult = this.Processor.SetValue(ImportResult, ObjectToBeFilled, BooleanProperty, Cell);

            cellProcessResult.Should().BeTrue();

            ImportResult.Errors.Should().BeNullOrEmpty();
            ObjectToBeFilled.BooleanColumn.Should().Be(false);
        }


        [TestMethod]
        public void Process_Letter_AsError()
        {
            Cell.Value = "AAAAA";
            var cellProcessResult = this.Processor.SetValue(ImportResult, ObjectToBeFilled, BooleanProperty, Cell);

            cellProcessResult.Should().BeFalse();
            ImportResult.Errors.Should().NotBeNullOrEmpty();
            ImportResult.Errors.Count.Should().Be(1);
            ImportResult.Errors[0].Column.Should().Be(1);
            ImportResult.Errors[0].Row.Should().Be(1);
            ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

        }

        [TestMethod]
        public void Process_EmptyString_AsNull()
        {
            Cell.Value = "";
            var cellProcessResult = this.Processor.SetValue(ImportResult, ObjectToBeFilled, BooleanProperty, Cell);

            cellProcessResult.Should().BeTrue();
            ObjectToBeFilled.NullableBooleanColumn.Should().Be(null);

        }

        [TestMethod]
        public void Process_Null_AsNull()
        {
            Cell.Value = null;
            var cellProcessResult = this.Processor.SetValue(ImportResult, ObjectToBeFilled, BooleanProperty, Cell);

            cellProcessResult.Should().BeTrue();
            ObjectToBeFilled.NullableBooleanColumn.Should().Be(null);
        }

        public IXLCell GetXLCell()
        {
            var wk = new XLWorkbook();
            wk.AddWorksheet("Test1");
            return wk.Worksheet(1).Cell(1, 1);
        }


    }
}
