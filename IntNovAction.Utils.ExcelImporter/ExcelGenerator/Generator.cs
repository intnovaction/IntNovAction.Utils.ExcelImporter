using ClosedXML.Excel;
using IntNovAction.Utils.ExcelImporter;
using IntNovAction.Utils.ExcelImporter.CellProcessors;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace IntNovAction.Utils.ExcelImporter.ExcelGenerator
{
    internal class Generator<TImportInfo>
    {
        public XLWorkbook GenerateExcel(List<FieldImportInfo<TImportInfo>> _fieldsInfo)
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.Worksheets.Add(typeof(TImportInfo).Name);

                var row = sheet.Row(Constants.FirstExcelRow);

                for (var i = 0; i < _fieldsInfo.Count; i++)
                {
                    var field = _fieldsInfo[i];
                    var cell = row.Cell(i + 1);

                    cell.Value = field.ColumnName;
                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                    cell.Style.Font.Bold = true;
                    cell.WorksheetColumn().AdjustToContents();
                }

                sheet.SheetView.FreezeRows(Constants.FirstExcelRow);

                return workbook;
            }
        }
    }
}
