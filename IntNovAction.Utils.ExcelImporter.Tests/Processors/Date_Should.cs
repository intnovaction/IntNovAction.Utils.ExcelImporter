using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelImporter.CellProcessors;
using IntNovAction.Utils.ExcelImporter;
using IntNovAction.Utils.ExcelImporter.Tests.SampleClasses;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Reflection;

namespace IntNovAction.Utils.ExcelImporter.Tests.Processors
{
    [TestClass]
    public class Date_Should
    {
        private readonly PropertyInfo DateProperty;

        private DateCellProcessor<SampleImportInto> Processor;
        private ImportResult<SampleImportInto> ImportResult;
        private SampleImportInto ObjectToBeFilled;
        private SampleImportInto ObjectToBeRead;

        public IXLCell Cell { get; private set; }

        public Date_Should()
        {
            DateProperty = typeof(SampleImportInto).GetProperty(nameof(SampleImportInto.DateColumn));
        }

        [TestInitialize()]
        public void Initializer()
        {
            this.Processor = new DateCellProcessor<SampleImportInto>(false);

            this.ImportResult = new ImportResult<SampleImportInto>();
            this.ObjectToBeFilled = new SampleImportInto();
            this.ObjectToBeRead = new SampleImportInto();
            this.Cell = GetXLCell();

        }


        [TestMethod]
        public void Process_Date_Ok()
        {
            var date = DateTime.Now;

            ObjectToBeRead.DateColumn = date;
            this.Processor.SetValueFromObjectToExcel(ObjectToBeRead, DateProperty, Cell);
            Cell.Value.ToString().Should().Be(ObjectToBeRead.DateColumn.ToString());

            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, DateProperty, Cell);

            cellProcessResult.Should().BeTrue();

            ImportResult.Errors.Should().BeNullOrEmpty();
            ObjectToBeFilled.DateColumn.Date.Should().Be(date.Date);
        }

        [TestMethod]
        public void Process_Invalid_Date_As_Error()
        {

            Cell.Value = "2018-33-11";
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, DateProperty, Cell);

            cellProcessResult.Should().BeFalse();
            ImportResult.Errors.Should().NotBeNullOrEmpty();
            ImportResult.Errors.Count.Should().Be(1);
            ImportResult.Errors[0].Column.Should().Be(1);
            ImportResult.Errors[0].Row.Should().Be(1);
            ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);
            
        }

        [TestMethod]
        public void Process_Letter_AsError()
        {
            Cell.Value = "S";
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, DateProperty, Cell);

            cellProcessResult.Should().BeFalse();
            ImportResult.Errors.Should().NotBeNullOrEmpty();
            ImportResult.Errors.Count.Should().Be(1);
            ImportResult.Errors[0].Column.Should().Be(1);
            ImportResult.Errors[0].Row.Should().Be(1);
            ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

        }

        [TestMethod]
        public void Process_EmptyString_AsError()
        {
            Cell.Value = "";
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, DateProperty, Cell);

            cellProcessResult.Should().BeFalse();
            ImportResult.Errors.Should().NotBeNullOrEmpty();
            ImportResult.Errors.Count.Should().Be(1);
            ImportResult.Errors[0].Column.Should().Be(1);
            ImportResult.Errors[0].Row.Should().Be(1);
            ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

        }

        [TestMethod]
        public void Process_Null_AsError()
        {
            Cell.Value = null;
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, DateProperty, Cell);

            cellProcessResult.Should().BeFalse();
            ImportResult.Errors.Should().NotBeNullOrEmpty();
            ImportResult.Errors.Count.Should().Be(1);
            ImportResult.Errors[0].Column.Should().Be(1);
            ImportResult.Errors[0].Row.Should().Be(1);
            ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

        }

        public IXLCell GetXLCell()
        {
            var wk = new XLWorkbook();
            wk.AddWorksheet("Test1");
            return wk.Worksheet(1).Cell(1, 1);
        }


    }
}
