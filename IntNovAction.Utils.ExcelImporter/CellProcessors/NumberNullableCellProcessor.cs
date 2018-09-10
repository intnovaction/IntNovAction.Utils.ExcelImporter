using System.Reflection;
using ClosedXML.Excel;
using IntNovAction.Utils.Importer;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    internal class NumberNullableCellProcessor<TPropType, TImportInto> : CellProcessorBase<TImportInto>
        where TImportInto : class
        where TPropType: struct
    {
        

        internal override void SetValue(ImportResult<TImportInto> results,
            TImportInto objectToFill,
            PropertyInfo property,
            IXLCell cell)
        {
            if (cell.IsEmpty() || string.IsNullOrEmpty(cell.GetValue<string>()))
            {
                property.SetValue(objectToFill, null);
                return;
            }

            if (cell.TryGetValue(out TPropType valor))
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
