using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelMapper
{
    public class Excel
    {

        private readonly Settings _settings = new Settings();
        private readonly string _filename;
        private IReadOnlyDictionary<Type, IEnumerable<object>> _items;
        private bool _readCompleted;

        public Excel(string fileName, Action<Settings> setup, bool readImmediately = false)
        {
            _filename = fileName;
            setup(_settings);
            if (readImmediately)
                ReadAll();
        }

        private void ReadAll()
        {
            using (var workbook = new XLWorkbook(_filename))
            {
                _items = _settings.ApplyToWorkbook(workbook);
                _readCompleted = true;
            }

        }

        public IEnumerable<T> Get<T>(bool reread = false)
        {
            if (!_readCompleted || reread)
                ReadAll();
            return _items[typeof(T)].Cast<T>();
        }
    }
}