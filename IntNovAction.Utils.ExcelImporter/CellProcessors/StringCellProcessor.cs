using System.Reflection;
using ClosedXML.Excel;
using IntNovAction.Utils.ExcelImporter;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    internal class StringCellProcessor<TImportInto> : CellProcessorBase<TImportInto>
        where TImportInto : class
    {
       

        internal override bool SetValue(ImportResult<TImportInto> results,
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


    }
}
