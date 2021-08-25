using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ClosedXML.Excel;
using IntNovAction.Utils.ExcelImporter.CellProcessors;
using IntNovAction.Utils.ExcelImporter.ExcelGenerator;

namespace IntNovAction.Utils.ExcelImporter
{
    /// <summary>
    /// The importer
    /// </summary>
    /// <typeparam name="TImportInto">The class where we will import into</typeparam>
    public class Importer<TImportInto>
        where TImportInto : class, new()
    {

        private Stream _excelStream;

        internal object SetRowIndex(Func<object, object> p)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// El numero de hoja dentro del workbool
        /// </summary>
        private int _excelSheet = 1;
        private string _excelSheetName = null;

        private readonly BooleanOptions _boolOptions;

        private ErrorStrategy _errorStrategy = ErrorStrategy.DoNotAddElement;

        private readonly List<FieldImportInfo<TImportInto>> _fieldsInfo;

        private int _initialRow = 1;
        private int _initialColumn = 1;
        private int _initialRowForData = 2;

        private DuplicatedColumnStrategy _duplicatedColumStrategy;
        private Expression<Func<TImportInto, int>> _rowIndexExpression;

        public Importer()
        {
            this._fieldsInfo = new List<FieldImportInfo<TImportInto>>();
            this._boolOptions = new BooleanOptions();
        }


        /// <summary>
        /// Sets the strings used to parse the boolean values
        /// </summary>
        /// <param name="trueStrings">Strings which will cause the field to be true</param>
        /// <param name="falseStrings">Strings which will cause the field to be true</param>
        /// <remarks>The default values of the importer are:
        /// <list type="bullet">
        /// <item>TRUE: </item>
        /// <item>TRUE: </item>
        /// </list>
        /// </remarks>
        /// <returns></returns>
        public Importer<TImportInto> SetBooleanOptions(IEnumerable<string> trueStrings = null, IEnumerable<string> falseStrings = null)
        {

            if (trueStrings != null)
            {
                this._boolOptions.TrueStrings.Clear();
                this._boolOptions.TrueStrings.AddRange(trueStrings.ToList());
            }

            if (falseStrings != null)
            {
                this._boolOptions.FalseStrings.Clear();
                this._boolOptions.FalseStrings.AddRange(falseStrings.ToList());
            }

            return this;
        }

        /// <summary>
        /// Sets the excel file the importer will use
        /// </summary>
        /// <param name="excelStream">The XLSX document stream</param>
        /// <returns></returns>
        public Importer<TImportInto> FromExcel(Stream excelStream)
        {
            this._excelStream = excelStream;

            return this;
        }

        public Importer<TImportInto> SetInitialCoordintates(int headerStartColumn = 1, int headerStartRow = 1)
        {
            this._initialRow = headerStartRow;
            this._initialRowForData = headerStartRow + 1;
            this._initialColumn = headerStartColumn;

            return this;
        }

        /// <summary>
        /// Generates an excel that can be used to import data.
        /// The excel is created with the fields configured for the importer
        /// </summary>
        /// <returns></returns>
        public Stream GenerateExcel()
        {
            var generator = new Generator<TImportInto>();

            return generator.GenerateExcel(this._fieldsInfo);
        }

        /// <summary>
        /// Sets the excel file the importer will use
        /// </summary>
        /// <param name="excelStream">The XLSX document stream</param>
        /// <param name="sheetIndex">The 1-based sheet index</param>
        /// <returns></returns>
        public Importer<TImportInto> FromExcel(Stream excelStream, int sheetIndex)
        {
            this._excelStream = excelStream;
            this._excelSheet = sheetIndex;
            return this;
        }

        /// <summary>
        /// Specifies an error handling Strategy
        /// </summary>
        /// <param name="strategy"></param>
        /// <returns></returns>
        public Importer<TImportInto> SetErrorStrategy(ErrorStrategy strategy)
        {
            this._errorStrategy = strategy;
            return this;
        }

        /// <summary>
        /// Sets the excel file the importer will use
        /// </summary>
        /// <param name="excelStream"></param>
        /// <param name="sheetName">The sheet name</param>
        /// <returns></returns>
        public Importer<TImportInto> FromExcel(Stream excelStream, string sheetName)
        {
            this._excelStream = excelStream;
            this._excelSheetName = sheetName;
            return this;
        }

        /// <summary>
        /// Perform the importation from the excel
        /// </summary>
        /// <returns></returns>
        public ImportResult<TImportInto> Import()
        {
            var results = new ImportResult<TImportInto>();

            if (this._excelStream == null)
            {
                throw new ExcelImportException($"No excel stream provided, please use the {nameof(this.FromExcel)} method");
            }
            // abrimos el excel
            var book = new XLWorkbook(this._excelStream);
            var sheet = GetDataSheet(book);

            var numdFilas = sheet.LastRowUsed().RowNumber();
            if (numdFilas == 0)
            {
                results.Errors.Add(new ImportErrorInfo()
                {
                    ErrorType = ImportErrorType.SheetEmpty
                });
                return results;
            }

            var canContinue = AnalyzeHeaders(sheet, this._fieldsInfo, results.Errors);
            if (!canContinue)
            {
                return results;
            }

            for (int cellRow = this._initialRowForData; cellRow <= numdFilas; cellRow++)
            {
                var target = new TImportInto();

                bool isRowOk = true;

                foreach (var colImportInfo in this._fieldsInfo.Where(fInfo => fInfo.ColumnNumber != 0))
                {
                    var property = colImportInfo.MemberExpr.Member as PropertyInfo;
                    var realTarget = GetTargetObject(target, colImportInfo.MemberExpr);

                    if (property != null)
                    {
                        var cell = sheet.Row(cellRow).Cell(colImportInfo.ColumnNumber);

                        var processor = GetProperPropertyProcessor(property.PropertyType);

                        isRowOk &= processor.SetValue(results, realTarget, property, cell);
                    }
                }

                if (this._rowIndexExpression != null)
                {
                    var property = Util<TImportInto, int>.GetMemberExpression(this._rowIndexExpression).Member as PropertyInfo;

                    property.SetValue(target, cellRow);
                }

                if (isRowOk || this._errorStrategy == ErrorStrategy.AddElement)
                {
                    results.ImportedItems.Add(target);
                }

            }

            return results;
        }

        private object GetTargetObject(TImportInto target, MemberExpression memberExpr)
        {
            var expressionAsString = memberExpr.ToString();
            var objectPath = expressionAsString.Substring(expressionAsString.IndexOf(".") + 1).Split('.');
            if (objectPath.Length == 1)
            {
                return target;
            }
            else
            {

                object tempTarget = target;
                for (var i = 0; i < objectPath.Length - 1; i++)
                {
                    PropertyInfo propertyToGet = tempTarget.GetType().GetProperty(objectPath[i]);
                    tempTarget = propertyToGet.GetValue(tempTarget);
                }
                return tempTarget;
            }
        }

        public Importer<TImportInto> SetDuplicatedColumnsStrategy(DuplicatedColumnStrategy duplicatedColumnStrategy)
        {
            this._duplicatedColumStrategy = duplicatedColumnStrategy;
            return this;
        }

        private IXLWorksheet GetDataSheet(XLWorkbook book)
        {
            if (!string.IsNullOrWhiteSpace(this._excelSheetName))
            {
                return book.Worksheets.Worksheet(this._excelSheetName);
            }
            return book.Worksheets.Worksheet(this._excelSheet);
        }

        internal CellProcessorBase<TImportInto> GetProperPropertyProcessor(Type propertyType)
        {
            if (propertyType.FullName == typeof(int).FullName)
            {
                return new NumberCellProcessor<int, TImportInto>();
            }
            else if (propertyType.FullName == typeof(int?).FullName)
            {
                return new NumberNullableCellProcessor<int, TImportInto>();
            }
            else if (propertyType.FullName == typeof(decimal).FullName)
            {
                return new NumberCellProcessor<decimal, TImportInto>();
            }
            else if (propertyType.FullName == typeof(decimal?).FullName)
            {
                return new NumberNullableCellProcessor<decimal, TImportInto>();
            }
            else if (propertyType.FullName == typeof(float).FullName)
            {
                return new NumberCellProcessor<float, TImportInto>();
            }
            else if (propertyType.FullName == typeof(float?).FullName)
            {
                return new NumberNullableCellProcessor<float, TImportInto>();
            }
            else if (propertyType.FullName == typeof(DateTime).FullName)
            {
                return new DateCellProcessor<TImportInto>(false);
            }
            else if (propertyType.FullName == typeof(DateTime?).FullName)
            {
                return new DateCellProcessor<TImportInto>(true);
            }
            else if (propertyType.FullName == typeof(string).FullName)
            {
                return new StringCellProcessor<TImportInto>();
            }
            else if (propertyType.FullName == typeof(Boolean).FullName)
            {
                return new BooleanCellProcessor<TImportInto>(false, this._boolOptions);
            }
            else if (propertyType.FullName == typeof(Boolean?).FullName)
            {
                return new BooleanCellProcessor<TImportInto>(true, this._boolOptions);
            }

            throw new NotImplementedException($"The processor for {propertyType.FullName} is not implemented");
        }




        private bool AnalyzeHeaders(IXLWorksheet sheet,
            List<FieldImportInfo<TImportInto>> fieldsInfo,
            List<ImportErrorInfo> errors)
        {

            bool hasDuplicated = false;

            var columnNames = new Dictionary<string, int>();

            var firstRow = sheet.Row(this._initialRow);

            var column = this._initialColumn;
            var lastColumn = firstRow.LastCellUsed().Address.ColumnNumber;
            while (column <= lastColumn)
            {
                var cell = firstRow.Cell(column);
                if (cell.TryGetValue<string>(out string header))
                {
                    if (columnNames.ContainsKey(header))
                    {
                        hasDuplicated = hasDuplicated | true;

                        errors.Add(new ImportErrorInfo()
                        {
                            ErrorType = ImportErrorType.DuplicatedColumn,
                            Column = cell.Address.ColumnNumber,
                            ColumnName = cell.Address.ColumnLetter,
                            CellValue = header,
                            Row = cell.Address.RowNumber
                        });
                        if (this._duplicatedColumStrategy == DuplicatedColumnStrategy.TakeLast)
                        {
                            columnNames[header] = column;
                        }
                    }
                    else
                    {
                        columnNames.Add(header, column);
                    }
                }
                column++;
            }

            if (hasDuplicated && this._duplicatedColumStrategy == DuplicatedColumnStrategy.RaiseError)
            {
                return false;
            }

            foreach (var fieldInfo in fieldsInfo)
            {

                foreach (var header in columnNames.Keys)
                {
                    if (string.Equals(header, fieldInfo.ColumnName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        fieldInfo.ColumnNumber = columnNames[header];
                        break;
                    }
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

            return true;
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

        /// <summary>
        /// Fill the property with the row index
        /// </summary>
        /// <param name="memberAccessor"></param>
        /// <returns></returns>
        public Importer<TImportInto> SetRowIndex(Expression<Func<TImportInto, int>> memberAccessor)
        {
            this._rowIndexExpression = memberAccessor;
            return this;
        }
    }
}
