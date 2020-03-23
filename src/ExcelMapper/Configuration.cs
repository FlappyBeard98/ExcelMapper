using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace ExcelMapper
{
    public abstract class Configuration
    {
        internal abstract Type Type { get; }

        internal abstract IEnumerable<object> ApplyToWorkSheet(IXLWorksheet xlWorksheet);
        
    }
}