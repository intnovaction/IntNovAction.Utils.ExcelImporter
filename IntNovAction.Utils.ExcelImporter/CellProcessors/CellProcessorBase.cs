using ClosedXML.Excel;
using IntNovAction.Utils.Importer;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    abstract class CellProcessorBase<TImportInto>
        where TImportInto: class
    {
        internal abstract void SetValue(ImportResult<TImportInto> results,
            TImportInto objectToFill,
            PropertyInfo property,
            IXLCell cell);

        protected void AddError(ImportResult<TImportInto> results, IXLCell cell)
        {
            results.Errors.Add(new ImportErrorInfo()
            {
                Column = cell.Address.ColumnNumber,
                Row = cell.Address.RowNumber,
                ErrorType = ImportErrorType.InvalidValue
            });
        }
    }
}
