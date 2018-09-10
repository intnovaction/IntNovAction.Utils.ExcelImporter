using System.Reflection;
using ClosedXML.Excel;
using IntNovAction.Utils.Importer;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    internal abstract class CellProcessorBase<TImportInto>
        where TImportInto : class
    {
        

        internal abstract bool SetValue(ImportResult<TImportInto> results,
            TImportInto objectToFill,
            PropertyInfo property,
            IXLCell cell);

        protected void AddInvalidValueError(ImportResult<TImportInto> results, IXLCell cell)
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
