using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text;

namespace IntNovAction.Utils.ExcelImporter.CellProcessors
{
    static class Util<TImportInto, TMember>
    {
        internal static MemberExpression GetMemberExpression(Expression<Func<TImportInto, TMember>> expression)
        {
            MemberExpression me;
            switch (expression.Body.NodeType)
            {
                case ExpressionType.Convert:
                case ExpressionType.ConvertChecked:
                    var ue = expression.Body as UnaryExpression;
                    me = (ue?.Operand) as MemberExpression;
                    break;
                default:
                    me = expression.Body as MemberExpression;
                    break;
            }

            return me;
        }
    }
}
