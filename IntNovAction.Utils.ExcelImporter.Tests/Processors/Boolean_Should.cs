﻿using ClosedXML.Excel;
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
    public class Boolean_Should
    {
        private readonly PropertyInfo BooleanProperty;

        private BooleanCellProcessor<SampleImportInto> Processor;
        private ImportResult<SampleImportInto> ImportResult;
        private SampleImportInto ObjectToBeFilled;
        private SampleImportInto ObjectToBeRead;

        public IXLCell Cell { get; private set; }

        public Boolean_Should()
        {
            BooleanProperty = typeof(SampleImportInto).GetProperty(nameof(SampleImportInto.BooleanColumn));
        }

        [TestInitialize()]
        public void Initializer()
        {
            var options = new BooleanOptions();

            this.Processor = new BooleanCellProcessor<SampleImportInto>(false, options);

            this.ImportResult = new ImportResult<SampleImportInto>();
            this.ObjectToBeFilled = new SampleImportInto();
            this.ObjectToBeRead = new SampleImportInto();
            this.Cell = GetXLCell();

        }


        [TestMethod]
        public void Process_Boolean_Alternate_Strings()
        {
            var options = new BooleanOptions();
            options.TrueStrings.Clear();
            options.TrueStrings.Add("sep");

            options.FalseStrings.Clear();
            options.FalseStrings.Add("nope");

            var altProcessor = new BooleanCellProcessor<SampleImportInto>(false, options);

            Cell.Value = "sep";
            var cellProcessResult = altProcessor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, BooleanProperty, Cell);
            ObjectToBeFilled.BooleanColumn.Should().Be(true);
            cellProcessResult.Should().BeTrue();

            Cell.Value = "nope";
            cellProcessResult = altProcessor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, BooleanProperty, Cell);
            ObjectToBeFilled.BooleanColumn.Should().Be(false);
            cellProcessResult.Should().BeTrue();

        }

        [TestMethod]
        public void Process_Boolean_True_Value()
        {
            ObjectToBeRead.BooleanColumn = true;
            this.Processor.SetValueFromObjectToExcel(ObjectToBeRead, BooleanProperty, Cell);
            Cell.Value.Should().Be(ObjectToBeRead.BooleanColumn);

            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, BooleanProperty, Cell);
            cellProcessResult.Should().BeTrue();

            ImportResult.Errors.Should().BeNullOrEmpty();
            ObjectToBeFilled.BooleanColumn.Should().Be(true);


        }

        [TestMethod]
        public void Process_Boolean_False_Value()
        {
            ObjectToBeRead.BooleanColumn = false;
            this.Processor.SetValueFromObjectToExcel(ObjectToBeRead, BooleanProperty, Cell);
            Cell.Value.Should().Be(ObjectToBeRead.BooleanColumn);

            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, BooleanProperty, Cell);

            cellProcessResult.Should().BeTrue();

            ImportResult.Errors.Should().BeNullOrEmpty();
            ObjectToBeFilled.BooleanColumn.Should().Be(false);
        }


        [TestMethod]
        public void Process_Letter_AsError()
        {
            Cell.Value = "AAAAA";
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, BooleanProperty, Cell);

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
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, BooleanProperty, Cell);

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
            var cellProcessResult = this.Processor.SetValueFromExcelToObject(ImportResult, ObjectToBeFilled, BooleanProperty, Cell);

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
