using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace ExcelMapper
{
    public abstract class Configuration
    {
        public abstract Type Type { get; }

        public abstract IEnumerable<object> ApplyToWorkSheet(IXLWorksheet xlWorksheet);
        
    }
}