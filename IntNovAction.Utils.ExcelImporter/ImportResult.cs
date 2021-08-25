using System.Collections.Generic;
using System.Linq;

namespace IntNovAction.Utils.ExcelImporter
{
    public class ImportResult<TImportInto>
        where TImportInto : class
    {

        public ImportResult()
        {
            this.ImportedItems = new List<TImportInto>();
            this.Errors = new List<ImportErrorInfo>();
        }

        public List<TImportInto> ImportedItems { get; set; }

        public ImportErrorResult Result
        {
            get
            {
                if (this.Errors.Count() == 0)
                {
                    return ImportErrorResult.Ok;
                }
                else if (this.Errors.Count() != 0 && this.ImportedItems.Count() != 0)
                {
                    return ImportErrorResult.PartialOk;
                }

                return ImportErrorResult.Error;

            }
        }

        public List<ImportErrorInfo> Errors { get; set; }
    }
}
