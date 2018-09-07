using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace IntNovAction.Utils.Importer
{
    /// <summary>
    /// The importer
    /// </summary>
    /// <typeparam name="TImportInto">The class where we will import into</typeparam>
    public class Importer<TImportInto>
        where TImportInto : class, new()
    {
        private Stream _excelStream;

        private readonly int _excelSheet = 1;
        private readonly List<FieldImportInfo<TImportInto>> _fieldsInfo;


        public Importer()
        {
            this._fieldsInfo = new List<Utils.Importer.FieldImportInfo<TImportInto>>();
        }

        /// <summary>
        /// Sets the excel file the importer will use
        /// </summary>
        /// <param name="excelPath"></param>
        /// <returns></returns>
        public Importer<TImportInto> FromExcel(Stream excelStream)
        {
            this._excelStream = excelStream;

            return this;
        }

        public ImportResult<TImportInto> Import()
        {
            var results = new ImportResult<TImportInto>();

            if (this._excelStream == null)
            {
                throw new ExcelImportException("No excel stream passed");
            }
            // abrimos el excel
            var book = new XLWorkbook(this._excelStream);
            var sheet = book.Worksheets.Worksheet(this._excelSheet);

            var numdFilas = sheet.LastRowUsed().RowNumber();
            if (numdFilas == 0)
            {
                results.Errors.Add(new ImportErrorInfo()
                {
                    ErrorType = ImportErrorType.SheetEmpty
                });
                return results;
            }

            AnalyzeHeaders(sheet, this._fieldsInfo, results.Errors);


            for (int cellRow = 2; cellRow <= numdFilas; cellRow++)
            {
                var imported = new TImportInto();

                foreach (var colImportInfo in this._fieldsInfo.Where(fInfo => fInfo.ColumnNumber != 0))
                {

                    var property = colImportInfo.MemberExpr.Member as PropertyInfo;

                    if (property != null)
                    {
                        var cell = sheet.Row(cellRow).Cell(colImportInfo.ColumnNumber);

                        string propertyTypeName = GetProperPropertyTypeName(property.PropertyType);

                        if (propertyTypeName == typeof(int).FullName)
                        {
                            SetIntegerValue(results, imported, property, cell);
                        }
                        else if (propertyTypeName == typeof(int?).FullName)
                        {
                            SetNullableIntegerValue(results, imported, property, cell);
                        }
                        else
                        {
                            throw new PropertyTypeNotSupportedException();
                        }



                        //    property.SetValue(imported, value, null);
                    }

                    results.ImportedItems.Add(imported);


                }

            }

            return results;
        }

        internal static string GetProperPropertyTypeName(Type propertyType)
        {
            return propertyType.FullName;
        }

        private static void SetIntegerValue(ImportResult<TImportInto> results, TImportInto objectToFill, PropertyInfo property, IXLCell cell)
        {
            if (cell.TryGetValue(out int valor))
            {
                property.SetValue(objectToFill, valor);
            }
            else
            {
                AddError(results, cell);
            }
        }

        private static void SetNullableIntegerValue(ImportResult<TImportInto> results, TImportInto objectToFill, PropertyInfo property, IXLCell cell)
        {
            if (String.IsNullOrWhiteSpace(cell.GetString()))
            {
                property.SetValue(objectToFill, null);
            }
            else
            {
                if (cell.TryGetValue(out int valor))
                {
                    property.SetValue(objectToFill, valor);
                }
                else
                {
                    AddError(results, cell);
                }
            }
        }

        private static void AddError(ImportResult<TImportInto> results, IXLCell cell)
        {
            results.Errors.Add(new ImportErrorInfo()
            {
                Column = cell.Address.ColumnNumber,
                Row = cell.Address.RowNumber,
                ErrorType = ImportErrorType.InvalidValue
            });
        }

        private static void AnalyzeHeaders(IXLWorksheet sheet,
            List<FieldImportInfo<TImportInto>> fieldsInfo,
            List<ImportErrorInfo> errors)
        {
            var firstRow = sheet.Row(1);
            foreach (var fieldInfo in fieldsInfo)
            {
                var column = 1;
                var lastColumn = firstRow.LastCellUsed().Address.ColumnNumber;
                while (column <= lastColumn)
                {
                    if (firstRow.Cell(column).TryGetValue<string>(out string header))
                    {
                        if (string.Equals(header, fieldInfo.ColumnName, StringComparison.CurrentCultureIgnoreCase))
                        {
                            fieldInfo.ColumnNumber = column;
                        }
                    }
                    column++;
                }

                if (fieldInfo.ColumnNumber == 0)
                {
                    errors.Add(new ImportErrorInfo()
                    {
                        ErrorType = ImportErrorType.ColumnNotFound,
                        ColumnName = fieldInfo.ColumnName
                    });
                }
            }
        }

        /// <summary>
        /// Maps a column of the excel to a property of <typeparamref name="TImportInto"/> 
        /// </summary>
        /// <param name="memberAccessor"></param>
        /// <param name="columnName"></param>
        public Importer<TImportInto> For(Expression<Func<TImportInto, object>> memberAccessor, string columnName)
        {
            var fInfo = new FieldImportInfo<TImportInto>(memberAccessor)
            {
                ColumnName = columnName
            };

            this._fieldsInfo.Add(fInfo);

            return this;
        }


    }
}
