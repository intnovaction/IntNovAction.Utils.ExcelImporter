using System;
using System.Linq.Expressions;
using IntNovAction.Utils.ExcelImporter.CellProcessors;

namespace IntNovAction.Utils.Importer
{
    internal class FieldImportInfo<TImportInto>
    {
        public FieldImportInfo(Expression<Func<TImportInto, dynamic>> expression)
        {
            var me = Util<TImportInto, dynamic>.GetMemberExpression(expression);
            this.MemberExpr = me;
        }

        

        public string ColumnName { get; set; }
        public int ColumnNumber { get; internal set; }
        public MemberExpression MemberExpr { get; }

    }
}