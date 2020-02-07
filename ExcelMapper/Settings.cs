using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelMapper
{
    public class Settings
    {
        private readonly Dictionary<string, Configuration> _configurations =
            new Dictionary<string, Configuration>();

        public void Add(Configuration configuration, string sheet)
        {
            sheet = string.IsNullOrWhiteSpace(sheet) ? configuration.Type.Name : sheet;
            _configurations.Add(sheet, configuration);
        }

        public IReadOnlyDictionary<Type, IEnumerable<object>> ApplyToWorkbook(IXLWorkbook workbook) =>
            workbook.Worksheets
                    .Select(p =>
                        _configurations.TryGetValue(p.Name, out var c)
                            ? new {Configuration = c, Sheet = p}
                            : null)
                    .Where(p => p != null)
                    .Select(p => (p.Configuration.Type,p.Configuration.ApplyToWorkSheet(p.Sheet)))
                    .ToDictionary(p => p.Type, p => p.Item2);

    }
}