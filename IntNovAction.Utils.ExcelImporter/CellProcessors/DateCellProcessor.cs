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

        internal override void SetValue(ImportResult<TImportInto> results,
            TImportInto objectToFill,
            PropertyInfo property,
            IXLCell cell)
        {

            if (cell.IsEmpty() || string.IsNullOrEmpty(cell.GetValue<string>()))
            {
                if (IsNullable)
                {
                    property.SetValue(objectToFill, null);
                }
                else
                {
                    base.AddInvalidValueError(results, cell);
                }
                return;
            }

            if (cell.TryGetValue(out DateTime valor))
            {
                property.SetValue(objectToFill, valor);
            }
            else
            {
                base.AddInvalidValueError(results, cell);
            }
        }



    }
}
