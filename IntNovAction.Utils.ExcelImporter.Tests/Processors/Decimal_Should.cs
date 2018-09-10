﻿using ClosedXML.Excel;
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
    public class Decimal_Should
    {
        private readonly PropertyInfo DecimalProperty;

        private NumberCellProcessor<decimal, SampleImportInto> Processor;
        private ImportResult<SampleImportInto> ImportResult;
        private SampleImportInto ObjectToBeFilled;

        public IXLCell Cell { get; private set; }

        public Decimal_Should()
        {
            DecimalProperty = typeof(SampleImportInto).GetProperty(nameof(SampleImportInto.DecimalColumn));
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

            Cell.Value = 1;
            this.Processor.SetValue(ImportResult, ObjectToBeFilled, DecimalProperty, Cell);

            ImportResult.Errors.Should().BeNullOrEmpty();
            ObjectToBeFilled.DecimalColumn.Should().Be(1);
        }

        [TestMethod]
        public void Process_Letter_AsError()
        {
            Cell.Value = "S";
            this.Processor.SetValue(ImportResult, ObjectToBeFilled, DecimalProperty, Cell);

            ImportResult.Errors.Should().NotBeNullOrEmpty();
            ImportResult.Errors.Count.Should().Be(1);
            ImportResult.Errors[0].Column.Should().Be(1);
            ImportResult.Errors[0].Row.Should().Be(1);
            ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

            ObjectToBeFilled.DecimalColumn.Should().Be(0);
        }

        [TestMethod]
        public void Process_EmptyString_AsError()
        {
            Cell.Value = "";
            this.Processor.SetValue(ImportResult, ObjectToBeFilled, DecimalProperty, Cell);

            ImportResult.Errors.Should().NotBeNullOrEmpty();
            ImportResult.Errors.Count.Should().Be(1);
            ImportResult.Errors[0].Column.Should().Be(1);
            ImportResult.Errors[0].Row.Should().Be(1);
            ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

            ObjectToBeFilled.DecimalColumn.Should().Be(0);
        }

        [TestMethod]
        public void Process_Null_AsError()
        {
            Cell.Value = null;
            this.Processor.SetValue(ImportResult, ObjectToBeFilled, DecimalProperty, Cell);

            ImportResult.Errors.Should().NotBeNullOrEmpty();
            ImportResult.Errors.Count.Should().Be(1);
            ImportResult.Errors[0].Column.Should().Be(1);
            ImportResult.Errors[0].Row.Should().Be(1);
            ImportResult.Errors[0].ErrorType.Should().Be(ImportErrorType.InvalidValue);

            ObjectToBeFilled.DecimalColumn.Should().Be(0);
        }

        public IXLCell GetXLCell()
        {
            var wk = new XLWorkbook();
            wk.AddWorksheet("Test1");
            return wk.Worksheet(1).Cell(1, 1);
        }


    }
}