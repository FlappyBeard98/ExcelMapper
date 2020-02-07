using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ClosedXML.Excel;

namespace ExcelMapper
{
    public class HeaderConfiguration<T> : Configuration where T:new()
    {
        public override  Type Type  => typeof(T);

        public override IEnumerable<object> ApplyToWorkSheet(IXLWorksheet xlWorksheet)
        {
            {
                var result = new List<object>();
                foreach (var row in xlWorksheet.Rows().Skip(1))
                {
                    if (row.IsEmpty())
                        break;
                    
                    var instance = Activator.CreateInstance(Type);
                    foreach (var column in xlWorksheet.Columns())
                    {
                    
                        var member =  _columnsConfiguration.TryGetValue(column.Cell(1).Value?.ToString() ?? "", out var m) ? m : null;
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


        private readonly Dictionary<string, MemberInfo> _columnsConfiguration = new Dictionary<string, MemberInfo>();

        public void Add(Expression<Func<T, object>> prop, string header)
        {
            var memberInfo = prop.GetMemberFromExpression();
            header = string.IsNullOrWhiteSpace(header) ? memberInfo.Name : header;
            _columnsConfiguration.Add(header, memberInfo);
        }
    }
}