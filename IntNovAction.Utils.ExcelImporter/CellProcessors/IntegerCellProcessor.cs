﻿using ClosedXML.Excel;
using IntNovAction.Utils.Importer;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    internal  class IntegerCellProcessor<TImportInto> : CellProcessorBase<TImportInto>
        where TImportInto: class
    {
        internal override void SetValue(ImportResult<TImportInto> results, 
            TImportInto objectToFill, 
            PropertyInfo property, 
            IXLCell cell)
        {
            if (cell.TryGetValue(out int valor))
            {
                property.SetValue(objectToFill, valor);
            }
            else
            {
                base.AddError(results, cell);
            }
        }

        
    }
}
