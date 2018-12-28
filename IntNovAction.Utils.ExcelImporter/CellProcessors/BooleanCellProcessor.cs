using System;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using IntNovAction.Utils.Importer;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    internal class BooleanCellProcessor<TImportInto> : CellProcessorBase<TImportInto>
        where TImportInto : class
    {

        private readonly bool IsNullable;
        private readonly BooleanOptions _options;

        public BooleanCellProcessor(bool nullable, BooleanOptions options) : base()
        {
            IsNullable = nullable;
            _options = options;
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

            var cellContent = cell.GetString();

            if (this._options.TrueStrings.Any(p => p.Equals(cellContent, StringComparison.CurrentCultureIgnoreCase)))
            {
                property.SetValue(objectToFill, true);
                return true;
            }

            if (this._options.FalseStrings.Any(p => p.Equals(cellContent, StringComparison.CurrentCultureIgnoreCase)))
            {
                property.SetValue(objectToFill, false);
                return true;
            }

            base.AddInvalidValueError(results, cell);
            return false;

        }



    }
}
