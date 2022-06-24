using System.Reflection;
using ClosedXML.Excel;
using IntNovAction.Utils.ExcelImporter;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    internal abstract class CellProcessorBase<TImportInto>
        where TImportInto : class
    {
        

        internal abstract bool SetValueFromExcelToObject(ImportResult<TImportInto> results,
            object target,
            PropertyInfo property,
            IXLCell cell);

        internal abstract bool SetValueFromObjectToExcel(object objectToRead,
            PropertyInfo property,
            IXLCell target);

        protected void AddInvalidValueError(ImportResult<TImportInto> results, IXLCell cell)
        {
            results.Errors.Add(new ImportErrorInfo()
            {
                Column = cell.Address.ColumnNumber,
                Row = cell.Address.RowNumber,
                ColumnName = cell.Address.ColumnLetter,
                CellValue = cell.GetString(),
                ErrorType = ImportErrorType.InvalidValue
            });
        }
    }
}
