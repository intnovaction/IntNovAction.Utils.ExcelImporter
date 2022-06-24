using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using FluentAssertions;
using IntNovAction.Utils.ExcelImporter.Tests.SampleClasses;
using IntNovAction.Utils.ExcelImporter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System;

namespace IntNovAction.Utils.ExcelImporter.Tests
{
    [TestClass]
    public class Importer_Should
    {
        [TestMethod]
        public void Import_FromExcel_SheetOk()
        {
            var importer = new Importer<SampleImportInto>();

            using (var stream = OpenExcel())
            {
                var lista = importer
                    .FromExcel(stream)
                    .For(p => p.IntColumn, "Int Column")
                    .For(p => p.FloatColumn, "Float Column")
                    .For(p => p.DecimalColumn, "Decimal Column")
                    .For(p => p.NullableIntColumn, "Nullable Int Column")
                    .For(p => p.NullableFloatColumn, "Nullable Float Column")
                    .For(p => p.NullableDecimalColumn, "Nullable Decimal Column")
                    .For(p => p.StringColumn, "String Column")
                    .For(p => p.DateColumn, "Date Column")
                    .For(p => p.NullableDateColumn, "Nullable Date Column")
                    .For(p => p.BooleanColumn, "Boolean Column")
                    .For(p => p.NullableBooleanColumn, "Nullable Boolean Column")
                    .Import();

                lista.Result.Should().Be(ImportErrorResult.Ok);

                lista.Errors.Should().NotBeNull();
                lista.Errors.Should().BeEmpty();

                lista.ImportedItems.Should().NotBeNull();
                lista.ImportedItems.Count().Should().Be(5);
            }
        }

        [TestMethod]
        public void Import_FromExcel_SheetError()
        {
            var importer = new Importer<SampleImportInto>();

            using (var stream = OpenExcel())
            {
                var lista = importer
                    .FromExcel(stream, "Data With Errors")
                    .For(p => p.IntColumn, "Int Column")
                    .For(p => p.FloatColumn, "Float Column")
                    .For(p => p.DecimalColumn, "Decimal Column")
                    .For(p => p.NullableIntColumn, "Nullable Int Column")
                    .For(p => p.NullableFloatColumn, "Nullable Float Column")
                    .For(p => p.NullableDecimalColumn, "Nullable Decimal Column")
                    .For(p => p.StringColumn, "String Column")
                    .For(p => p.DateColumn, "Date Column")
                    .For(p => p.NullableDateColumn, "Nullable Date Column")
                    .For(p => p.BooleanColumn, "Boolean Column")
                    .For(p => p.NullableBooleanColumn, "Nullable Boolean Column")
                    .Import();

                lista.Result.Should().Be(ImportErrorResult.PartialOk);

                lista.Errors.Should().NotBeNullOrEmpty();

                lista.ImportedItems.Should().NotBeNull();
                lista.ImportedItems.Count().Should().Be(3);
            }
        }

        [TestMethod]
        public void Import_FromExcel_SheetError_AddAll()
        {
            var importer = new Importer<SampleImportInto>();

            using (var stream = OpenExcel())
            {
                var lista = importer
                    .FromExcel(stream, "Data With Errors")
                    .SetErrorStrategy(ErrorStrategy.AddElement)
                    .For(p => p.IntColumn, "Int Column")
                    .For(p => p.FloatColumn, "Float Column")
                    .For(p => p.DecimalColumn, "Decimal Column")
                    .For(p => p.NullableIntColumn, "Nullable Int Column")
                    .For(p => p.NullableFloatColumn, "Nullable Float Column")
                    .For(p => p.NullableDecimalColumn, "Nullable Decimal Column")
                    .For(p => p.StringColumn, "String Column")
                    .For(p => p.DateColumn, "Date Column")
                    .For(p => p.NullableDateColumn, "Nullable Date Column")
                    .For(p => p.BooleanColumn, "Boolean Column")
                    .For(p => p.NullableBooleanColumn, "Nullable Boolean Column")
                    .Import();

                lista.Result.Should().Be(ImportErrorResult.PartialOk);

                lista.Errors.Should().NotBeNullOrEmpty();

                lista.ImportedItems.Should().NotBeNull();
                lista.ImportedItems.Count().Should().Be(5);
            }
        }

        public Stream OpenExcel()
        {
            var stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("IntNovAction.Utils.ExcelImporter.Tests.SampleExcels.SampleExcel.xlsx");

            return stream;
        }


        [TestMethod]
        public void Show_Error_When_Columns_Are_Duplicated()
        {
            var importer = new Importer<SampleImportInto>();

            using (var stream = OpenExcel())
            {
                var lista = importer
                    .SetDuplicatedColumnsStrategy(DuplicatedColumnStrategy.RaiseError)
                    .FromExcel(stream, "Duplicated Columns")
                    .SetErrorStrategy(ErrorStrategy.AddElement)
                    .For(p => p.IntColumn, "Int Column")
                    .Import();

                lista.Result.Should().Be(ImportErrorResult.Error);

                lista.Errors.Should().NotBeNullOrEmpty();

                lista.ImportedItems.Should().NotBeNull();
                lista.ImportedItems.Should().BeEmpty();
            }
        }

        [TestMethod]
        public void Take_First_Value_When_Columns_Are_Duplicated_And_Strategy_Set()
        {
            var importer = new Importer<SampleImportInto>();

            using (var stream = OpenExcel())
            {
                var lista = importer
                    .SetDuplicatedColumnsStrategy(DuplicatedColumnStrategy.TakeFirst)
                    .FromExcel(stream, "Duplicated Columns")
                    .SetErrorStrategy(ErrorStrategy.AddElement)
                    .For(p => p.IntColumn, "Int Column")
                    .Import();

                lista.Result.Should().Be(ImportErrorResult.PartialOk);

                lista.Errors.Should().NotBeNullOrEmpty();

                lista.ImportedItems.Should().NotBeNullOrEmpty();
                lista.ImportedItems[0].IntColumn.Should().Be(1);
            }
        }

        [TestMethod]
        public void Take_Last_Value_When_Columns_Are_Duplicated_And_Strategy_Set()
        {
            var importer = new Importer<SampleImportInto>();

            using (var stream = OpenExcel())
            {
                var lista = importer
                    .SetDuplicatedColumnsStrategy(DuplicatedColumnStrategy.TakeLast)
                    .FromExcel(stream, "Duplicated Columns")
                    .SetErrorStrategy(ErrorStrategy.AddElement)
                    .For(p => p.IntColumn, "Int Column")
                    .Import();

                lista.Result.Should().Be(ImportErrorResult.PartialOk);

                lista.Errors.Should().NotBeNullOrEmpty();

                lista.ImportedItems.Should().NotBeNullOrEmpty();
                lista.ImportedItems[0].IntColumn.Should().Be(33);
            }
        }

        [TestMethod]
        public void Fill_RowIndex_Property()
        {
            var importer = new Importer<SampleImportInto>();

            using (var stream = OpenExcel())
            {
                var lista = importer
                    .FromExcel(stream)
                    .SetRowIndex(p => p.RowIndex)
                    .For(p => p.IntColumn, "Int Column")
                    .Import();

                lista.Result.Should().Be(ImportErrorResult.Ok);

                lista.Errors.Should().NotBeNull();
                lista.Errors.Should().BeEmpty();

                lista.ImportedItems.Should().NotBeNullOrEmpty();
                lista.ImportedItems[0].RowIndex = 1;
                lista.ImportedItems[4].RowIndex = 5;
            }
        }


        [TestMethod]
        public void Generate_Excel_From_Importer()
        {
            var importer = new Importer<SampleImportInto>();

            importer
                .For(p => p.NullableIntColumn, "Nullable Int Column")
                .For(p => p.BooleanColumn, "Bool Column")
                .For(p => p.DateColumn, "Date column")
                .For(p => p.DecimalColumn, "Decimal Column");

            using (var excelStream = importer.GenerateExcel())
            {
                excelStream.Should().NotBeNull();

                var book = new XLWorkbook(excelStream);
                book.Should().NotBeNull();
                book.Worksheets.Count().Should().Be(1);

                var worksheet = book.Worksheet(1);
                worksheet.Name.Should().Be("SampleImportInto");

                worksheet.Row(1).Cell(1).Value.ToString().Should().Be("Nullable Int Column");
                worksheet.Row(1).Cell(2).Value.ToString().Should().Be("Bool Column");
                worksheet.Row(1).Cell(3).Value.ToString().Should().Be("Date column");
                worksheet.Row(1).Cell(4).Value.ToString().Should().Be("Decimal Column");
            }

        }

        [TestMethod]
        public void Generate_Excel_From_Importer_With_SampleData()
        {
            var importer = new Importer<SampleImportInto>();

            importer
                .For(p => p.NullableIntColumn, "Nullable Int Column")
                .For(p => p.BooleanColumn, "Bool Column")
                .For(p => p.DateColumn, "Date column")
                .For(p => p.DecimalColumn, "Decimal Column")
                .For(p => p.StringColumn, "String Column");

            var sampleData = new List<SampleImportInto>();
            var sampleDataItem1 = new SampleImportInto()
            {
                NullableIntColumn = 25,
                BooleanColumn = false,
                DateColumn = DateTime.Now,
                DecimalColumn = 2.23M,
                StringColumn = "test 1"
            };
            var sampleDataItem2 = new SampleImportInto()
            {
                NullableIntColumn = null,
                BooleanColumn = true,
                DateColumn = DateTime.Now.AddDays(-5),
                DecimalColumn = 2.23M,
                StringColumn = null
            };
            sampleData.Add(sampleDataItem1);
            sampleData.Add(sampleDataItem2);

            using (var excelStream = importer.GenerateExcel(sampleData))
            {
                excelStream.Should().NotBeNull();

                var book = new XLWorkbook(excelStream);
                book.Should().NotBeNull();
                book.Worksheets.Count().Should().Be(1);

                var worksheet = book.Worksheet(1);
                worksheet.Name.Should().Be("SampleImportInto");

                worksheet.Rows().Count().Should().Be(3);

                worksheet.Row(2).Cell(1).Value.ToString().Should().Be(sampleDataItem1.NullableIntColumn.ToString());
                worksheet.Row(2).Cell(2).Value.ToString().Should().Be(sampleDataItem1.BooleanColumn.ToString());
                worksheet.Row(2).Cell(3).Value.ToString().Should().Be(sampleDataItem1.DateColumn.ToString());
                worksheet.Row(2).Cell(4).Value.ToString().Should().Be(sampleDataItem1.DecimalColumn.ToString());
                worksheet.Row(2).Cell(5).Value.ToString().Should().Be(sampleDataItem1.StringColumn.ToString());

                worksheet.Row(3).Cell(1).Value.ToString().Should().Be(sampleDataItem2.NullableIntColumn.ToString());
                worksheet.Row(3).Cell(2).Value.ToString().Should().Be(sampleDataItem2.BooleanColumn.ToString());
                worksheet.Row(3).Cell(3).Value.ToString().Should().Be(sampleDataItem2.DateColumn.ToString());
                worksheet.Row(3).Cell(4).Value.ToString().Should().Be(sampleDataItem2.DecimalColumn.ToString());
                worksheet.Row(3).Cell(5).Value.ToString().Should().Be(string.Empty);
            }
        }


        [TestMethod]
        public void Import_From_Generated_Excel()
        {
            var importer = new Importer<SampleImportInto>();

            importer
                .For(p => p.NullableIntColumn, "Nullable Int Column")
                .For(p => p.BooleanColumn, "Bool Column")
                .For(p => p.DateColumn, "Date column")
                .For(p => p.DecimalColumn, "Decimal Column");

            XLWorkbook excelBook;

            using (var excelStream = importer.GenerateExcel())
            {
                excelStream.Should().NotBeNull();
                excelBook = new XLWorkbook(excelStream);

                var worksheet = excelBook.Worksheet(1);

                worksheet.Row(2).Cell(1).Value = 1;
                worksheet.Row(2).Cell(2).Value = 1;
                worksheet.Row(2).Cell(3).Value = "2018/1/1";
                worksheet.Row(2).Cell(4).Value = 15.2m;

                using (var destinationExcelStream = new MemoryStream())
                {
                    excelBook.SaveAs(destinationExcelStream);

                    var importResult = importer.FromExcel(destinationExcelStream)
                        .Import();

                    importResult.Should().NotBeNull();
                    importResult.Result.Should().NotBeNull();
                    importResult.ImportedItems.Should().NotBeNullOrEmpty();
                    importResult.ImportedItems.Count.Should().Be(1);

                    importResult.ImportedItems[0].NullableIntColumn.Should().Be(1);
                    importResult.ImportedItems[0].BooleanColumn.Should().Be(true);
                    importResult.ImportedItems[0].DateColumn.Should().Be(new System.DateTime(2018, 1, 1));
                    importResult.ImportedItems[0].DecimalColumn.Should().Be(15.2m);
                }

            }
        }


        [TestMethod]
        public void Generate_With_InnerClass()
        {
            var importer = new Importer<ClassWithInnerClass>();

            importer
                .For(p => p.Inner.PropInt, "Inner Column")
                .For(p => p.Inner, "Base Column");

            using (var stream = importer.GenerateExcel())
            {
                stream.Should().NotBeNull();

                var book = new XLWorkbook(stream);
                book.Should().NotBeNull();
                book.Worksheets.Count().Should().Be(1);

                var worksheet = book.Worksheet(1);

                worksheet.Row(1).Cell(1).Value.Should().Be("Inner Column");
                worksheet.Row(1).Cell(2).Value.Should().Be("Base Column");
            }
        }

        [TestMethod]
        public void Import_From_Generated_Excel_AnidatedClass()
        {
            var importer = new Importer<ClassWithInnerClass>();

            importer
                .For(p => p.Inner.PropInt, "Inner Column")
                .For(p => p.TestInt, "Base Column");

            XLWorkbook excelBook;

            using (var excelStream = importer.GenerateExcel())
            {
                excelStream.Should().NotBeNull();
                excelBook = new XLWorkbook(excelStream);

                var worksheet = excelBook.Worksheet(1);

                worksheet.Row(2).Cell(1).Value = 1;
                worksheet.Row(2).Cell(2).Value = 2;

                using (var destinationExcelStream = new MemoryStream())
                {
                    excelBook.SaveAs(destinationExcelStream);

                    var importResult = importer.FromExcel(destinationExcelStream)
                        .Import();

                    importResult.Should().NotBeNull();
                    importResult.Result.Should().NotBeNull();
                    importResult.ImportedItems.Should().NotBeNullOrEmpty();
                    importResult.ImportedItems.Count.Should().Be(1);

                    importResult.ImportedItems[0].Inner.PropInt.Should().Be(1);
                    importResult.ImportedItems[0].TestInt.Should().Be(2);
                }

            }
        }
    }
}
