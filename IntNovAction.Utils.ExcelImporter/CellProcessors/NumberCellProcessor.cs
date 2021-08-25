using System;
using System.Reflection;
using ClosedXML.Excel;
using IntNovAction.Utils.ExcelImporter;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    internal class NumberCellProcessor<TPropType, TImportInto> : CellProcessorBase<TImportInto>
        where TImportInto : class
        where TPropType : struct
    {
        

        internal override bool SetValue(ImportResult<TImportInto> results,
            object objectToFill,
            PropertyInfo property,
            IXLCell cell)
        {

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



    }
}
