using ClosedXML.Excel;
using IntNovAction.Utils.ExcelImporter;
using System.Collections.Generic;
using System.IO;

namespace IntNovAction.Utils.ExcelImporter.ExcelGenerator
{
    internal class Generator<TImportInfo>
    {
        public Stream GenerateExcel(List<FieldImportInfo<TImportInfo>> _fieldsInfo)
        {
            using (var workbook = new XLWorkbook())
            {

                var sheet = workbook.Worksheets.Add(typeof(TImportInfo).Name);

                var row = sheet.Row(1);

                for (var i = 0; i < _fieldsInfo.Count; i++)
                {
                    var field = _fieldsInfo[i];
                    var cell = row.Cell(i + 1);

                    cell.Value = field.ColumnName;
                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                    cell.Style.Font.Bold = true;
                    cell.WorksheetColumn().AdjustToContents();
                }

                sheet.SheetView.FreezeRows(1);

                var mStream = new MemoryStream();
                workbook.SaveAs(mStream);
                return mStream;

            }
        }
    }
}
