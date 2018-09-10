﻿using ClosedXML.Excel;
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
    public class Date_Should
    {
        private readonly PropertyInfo DateProperty;

        private DateCellProcessor<SampleImportInto> Processor;
        private ImportResult<SampleImportInto> ImportResult;
        private SampleImportInto ObjectToBeFilled;

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
            this.Cell = GetXLCell();

        }


        [TestMethod]
        public void Process_Date_Ok()
        {

            Cell.Value = "2018-01-01";
            this.Processor.SetValue(ImportResult, ObjectToBeFilled, DateProperty, Cell);

            ImportResult.Errors.Should().BeNullOrEmpty();
            ObjectToBeFilled.DateColumn.Should().Be(new DateTime(2018,1,1));
        }

        [TestMethod]
        public void Process_Invalid_Date_As_Error()
        {

            Cell.Value = "2018-33-11";
            this.Processor.SetValue(ImportResult, ObjectToBeFilled, DateProperty, Cell);

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
            this.Processor.SetValue(ImportResult, ObjectToBeFilled, DateProperty, Cell);

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
            this.Processor.SetValue(ImportResult, ObjectToBeFilled, DateProperty, Cell);

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
            this.Processor.SetValue(ImportResult, ObjectToBeFilled, DateProperty, Cell);

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