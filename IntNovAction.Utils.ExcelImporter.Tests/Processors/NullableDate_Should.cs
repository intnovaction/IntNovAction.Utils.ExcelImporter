using System;
using System.Reflection;
using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelImporter.CellProcessors;
using IntNovAction.Utils.Importer;
using IntNovAction.Utils.Importer.Tests.SampleClasses;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IntNovAction.Utils.ExcelImporter.Tests.Processors
{
    [TestClass]
    public class NullableDateTime_Should
    {
        private readonly PropertyInfo NullableDateProperty;

        private DateCellProcessor<SampleImportInto> Processor;
        private ImportResult<SampleImportInto> ImportResult;
        private SampleImportInto ObjectToBeFilled;

        public IXLCell Cell { get; private set; }

        public NullableDateTime_Should()
        {
            this.NullableDateProperty = typeof(SampleImportInto).GetProperty(nameof(SampleImportInto.NullableDateColumn));
        }

        [TestInitialize()]
        public void Initializer()
        {
            this.Processor = new DateCellProcessor<SampleImportInto>(true);

            this.ImportResult = new ImportResult<SampleImportInto>();
            this.ObjectToBeFilled = new SampleImportInto();
            this.Cell = GetXLCell();

        }


        [TestMethod]
        public void Process_Date_Ok()
        {

            this.Cell.Value = "2018-01-01";
            var cellProcessResult = this.Processor.SetValue(this.ImportResult, this.ObjectToBeFilled, this.NullableDateProperty, this.Cell);

            cellProcessResult.Should().BeTrue();

            this.ImportResult.Errors.Should().BeNullOrEmpty();
            this.ObjectToBeFilled.NullableDateColumn.Should().Be(new DateTime(2018, 1, 1));
        }

        [TestMethod]
        public void Process_Invalid_Date_As_Error()
        {

            this.Cell.Value = "2018-33-11";
            var cellProcessResult = this.Processor.SetValue(this.ImportResult, this.ObjectToBeFilled, this.NullableDateProperty, this.Cell);

            cellProcessResult.Should().BeFalse();
            this.ImportResult.Errors.Should().NotBeNullOrEmpty();
            this.ImportResult.Errors.Count.Should().Be(1);
            this.ImportResult.Errors[0].Column.Should().Be(1);
            this.ImportResult.Errors[0].Row.Should().Be(1);
            this.ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

        }

        [TestMethod]
        public void Process_Letter_AsError()
        {
            this.Cell.Value = "S";
            var cellProcessResult = this.Processor.SetValue(this.ImportResult, this.ObjectToBeFilled, this.NullableDateProperty, this.Cell);

            cellProcessResult.Should().BeFalse();
            this.ImportResult.Errors.Should().NotBeNullOrEmpty();
            this.ImportResult.Errors.Count.Should().Be(1);
            this.ImportResult.Errors[0].Column.Should().Be(1);
            this.ImportResult.Errors[0].Row.Should().Be(1);
            this.ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

        }

        [TestMethod]
        public void Process_EmptyString_AsNull()
        {
            this.Cell.Value = "";
            var cellProcessResult = this.Processor.SetValue(this.ImportResult, this.ObjectToBeFilled, this.NullableDateProperty, this.Cell);

            cellProcessResult.Should().BeTrue();

            this.ImportResult.Errors.Should().BeEmpty();

            this.ObjectToBeFilled.NullableDateColumn.Should().Be(null);
        }

        [TestMethod]
        public void Process_Null_AsNull()
        {
            this.Cell.Value = null;
            var cellProcessResult = this.Processor.SetValue(this.ImportResult, this.ObjectToBeFilled, this.NullableDateProperty, this.Cell);

            cellProcessResult.Should().BeTrue();
            this.ImportResult.Errors.Should().BeEmpty();

            this.ObjectToBeFilled.NullableDateColumn.Should().Be(null);
        }

        public IXLCell GetXLCell()
        {
            var wk = new XLWorkbook();
            wk.AddWorksheet("Test1");
            return wk.Worksheet(1).Cell(1, 1);
        }


    }
}
