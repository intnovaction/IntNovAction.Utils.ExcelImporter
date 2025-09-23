using System;
using System.Reflection;
using ClosedXML.Excel;
using IntNovAction.Utils.ExcelImporter;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    internal class DateCellProcessor<TImportInto> : CellProcessorBase<TImportInto>
        where TImportInto : class
    {
        private readonly bool IsNullable;

        public DateCellProcessor(bool nullable) : base()
        {
            IsNullable = nullable;
        }

        internal override bool SetValueFromExcelToObject(ImportResult<TImportInto> results,
            object objectToFill,
            PropertyInfo property,
            IXLCell cell)
        {

            if (cell.IsEmpty() || string.IsNullOrEmpty(cell.GetValue<string>()))
            {
                if (IsNullable)
                {
                    property.SetValue(objectToFill, null);
                    return true;
                }
                else
                {
                    base.AddInvalidValueError(results, cell);
                    return false;
                }
                
            }

            if (cell.TryGetValue(out DateTime valor))
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
                if (IsNullable)
                {
                    cellToFill.SetValue<string>(null);
                    return true;
                }
                else
                {
                    return false;
                }
            }

            var value = property.GetValue(objectToRead) as DateTime?;
            cellToFill.SetValue(value);

            return true;
        }
    }
}
