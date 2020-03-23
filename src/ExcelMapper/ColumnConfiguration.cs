using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using ClosedXML.Excel;

namespace ExcelMapper
{
    public class ColumnConfiguration<T> : Configuration  where T:new()
    {
        internal override  Type Type  => typeof(T);

        internal override IEnumerable<object> ApplyToWorkSheet(IXLWorksheet xlWorksheet)
        {
            {
                var result = new List<object>();
                foreach (var row in xlWorksheet.Rows())
                { 
                    if (row.IsEmpty())
                        break;
                    
                    var instance = Activator.CreateInstance(Type);
                    foreach (var column in xlWorksheet.Columns())
                    {

                        var member = _columnsConfiguration.TryGetValue(column.ColumnNumber(), out var m) ? m : null;
                        if (member == null)
                            continue;
                        var value = column.Cell(row.RowNumber()).Value;
                    
                        instance.Set(member,value);
                    }

                    result.Add(instance);
                }
                return result;
            }
        }

        private readonly Dictionary<int, MemberInfo> _columnsConfiguration = new Dictionary<int, MemberInfo>();

        internal void Add(Expression<Func<T, object>> prop, int column)
        {
            var memberInfo = prop.GetMemberFromExpression();
            _columnsConfiguration.Add(column, memberInfo);
        }
    }
}