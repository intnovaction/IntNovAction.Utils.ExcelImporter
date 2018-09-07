using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

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


            for (int i = 1; i < numdFilas; i++)
            {
                var imported = new TImportInto();

                foreach (var colImportInfo in this._fieldsInfo)
                {
                    var row = sheet.Row(i);

                }

            }


            return results;
        }


        private static void AnalyzeHeaders(IXLWorksheet sheet,
            List<FieldImportInfo<TImportInto>> fieldsInfo,
            List<ImportErrorInfo> errors)
        {
            var firstRow = sheet.Row(1);
            foreach (var fieldInfo in fieldsInfo)
            {
                var column = 0;
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

                if (column > lastColumn)
                {
                    errors.Add(new ImportErrorInfo()
                    {

                    });
                }
            }
        }

        /// <summary>
        /// Maps a column of the excel to a property of <typeparamref name="TImportInto"/> 
        /// </summary>
        /// <param name="memberAccessor"></param>
        /// <param name="columnName"></param>
        public Importer<TImportInto> For(Func<TImportInto, object> memberAccessor, string columnName)
        {
            var fInfo = new FieldImportInfo<TImportInto>()
            {
                Accessor = memberAccessor,
                ColumnName = columnName
            };

            this._fieldsInfo.Add(fInfo);

            return this;
        }
    }
}
