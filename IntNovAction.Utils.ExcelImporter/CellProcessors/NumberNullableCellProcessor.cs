using System.Reflection;
using ClosedXML.Excel;
//using IntNovAction.Utils.ExcelImporter;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    internal class NumberNullableCellProcessor<TPropType, TImportInto> : CellProcessorBase<TImportInto>
        where TImportInto : class
        where TPropType: struct
    {
        

        internal override bool SetValueFromExcelToObject(ImportResult<TImportInto> results,
            object objectToFill,
            PropertyInfo property,
            IXLCell cell)
        {
            if (cell.IsEmpty() || string.IsNullOrEmpty(cell.GetValue<string>()))
            {
                property.SetValue(objectToFill, null);
                return true;
            }

            if (cell.TryGetValue(out TPropType valor))
            {
                property.SetValue(objectToFill, valor);
                return true;
            }
            else
            {
                base.AddInvalidValueError(results, cell);
                return false;
            }
        }

        internal override bool SetValueFromObjectToExcel(object objectToRead,
            PropertyInfo property,
            IXLCell cellToFill)
        {
            if (property.GetValue(objectToRead) == null)
            {
                return true;
            }

            if (int.TryParse(property.GetValue(objectToRead).ToString(), out int resultInt))
            {
                cellToFill.SetValue(resultInt);
                return true;
            }
            else if (decimal.TryParse(property.GetValue(objectToRead).ToString(), out decimal resultDecimal))
            {
                cellToFill.SetValue(resultDecimal);
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
