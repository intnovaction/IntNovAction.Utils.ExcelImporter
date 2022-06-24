using System.Reflection;
using ClosedXML.Excel;
using IntNovAction.Utils.ExcelImporter;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    internal class StringCellProcessor<TImportInto> : CellProcessorBase<TImportInto>
        where TImportInto : class
    {
       

        internal override bool SetValueFromExcelToObject(ImportResult<TImportInto> results,
            object objectToFill,
            PropertyInfo property,
            IXLCell cell)
        {
            if (cell.IsEmpty())
            {
                property.SetValue(objectToFill, null);
            }
            else
            {
                property.SetValue(objectToFill, cell.GetString());
            }

            return true;

        }

        internal override bool SetValueFromObjectToExcel(object objectToRead,
            PropertyInfo property,
            IXLCell cellToFill)
        {
            if (property.GetValue(objectToRead) == null)
            {
                return true;
            }

            cellToFill.SetValue<string>(property.GetValue(objectToRead).ToString());
            return true;
        }
    }
}
