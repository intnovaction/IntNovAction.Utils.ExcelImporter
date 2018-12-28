using System;
using System.Reflection;
using ClosedXML.Excel;
using IntNovAction.Utils.Importer;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    internal class DateCellProcessor<TImportInto> : CellProcessorBase<TImportInto>
        where TImportInto : class
    {
        private bool IsNullable;

        public DateCellProcessor(bool nullable) : base()
        {
            IsNullable = nullable;
        }

        internal override bool SetValue(ImportResult<TImportInto> results,
            TImportInto objectToFill,
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



    }
}
