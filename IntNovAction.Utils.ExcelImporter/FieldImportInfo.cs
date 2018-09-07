using System;
using System.Linq.Expressions;

namespace IntNovAction.Utils.Importer
{
    internal class FieldImportInfo<TImportInto>
    {
        public FieldImportInfo(Expression<Func<TImportInto, dynamic>> expression)
        {
            MemberExpression me;
            switch (expression.Body.NodeType)
            {
                case ExpressionType.Convert:
                case ExpressionType.ConvertChecked:
                    var ue = expression.Body as UnaryExpression;
                    me = ((ue != null) ? ue.Operand : null) as MemberExpression;
                    break;
                default:
                    me = expression.Body as MemberExpression;
                    break;
            }
            this.MemberExpr = me;
        }

        
        public string ColumnName { get; set; }
        public int ColumnNumber { get; internal set; }
        public MemberExpression MemberExpr { get; }
    }
}