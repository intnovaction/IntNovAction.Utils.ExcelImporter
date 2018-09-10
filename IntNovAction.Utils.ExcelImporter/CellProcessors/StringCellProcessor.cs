using System.Reflection;
using ClosedXML.Excel;
using IntNovAction.Utils.Importer;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    internal class StringCellProcessor<TImportInto> : CellProcessorBase<TImportInto>
        where TImportInto : class
    {
       

        internal override void SetValue(ImportResult<TImportInto> results,
            TImportInto objectToFill,
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
        }


    }
}
