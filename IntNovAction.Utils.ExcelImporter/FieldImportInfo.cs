using System;

namespace IntNovAction.Utils.Importer
{
    internal class FieldImportInfo<TImportInto>
    {
        public FieldImportInfo()
        {
        }

        public Func<TImportInto, object> Accessor { get; set; }

        public string ColumnName { get; set; }
        public int ColumnNumber { get; internal set; }
    }
}