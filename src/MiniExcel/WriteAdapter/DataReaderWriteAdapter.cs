using MiniExcelLibs.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace MiniExcelLibs.WriteAdapter
{
    internal class DataReaderWriteAdapter : IMiniExcelWriteAdapter
    {
        private readonly IDataReader _reader;
        private readonly Configuration _configuration;

        public DataReaderWriteAdapter(IDataReader reader, Configuration configuration)
        {
            _reader = reader;
            _configuration = configuration;
        }

        public bool TryGetNonEnumeratedCount(out int count)
        {
            count = 0;
            return false;
        }

        public List<ExcelColumnInfo> GetColumns()
        {
            var props = new List<ExcelColumnInfo>();
            for (var i = 0; i < _reader.FieldCount; i++)
            {
                var columnName = _reader.GetName(i);

                if (!_configuration.DynamicColumnFirst)
                {
                    var prop = CustomPropertyHelper.GetColumnInfosFromDynamicConfiguration(columnName, _configuration);
                    props.Add(prop);
                    continue;
                }

                if (_configuration
                    .DynamicColumns
                    .Any(a => string.Equals(
                        a.Key,
                        columnName,
                        StringComparison.OrdinalIgnoreCase)))

                {
                    var prop = CustomPropertyHelper.GetColumnInfosFromDynamicConfiguration(columnName, _configuration);
                    props.Add(prop);
                }
            }
            return props;
        }

        public IEnumerable<IEnumerable<CellWriteInfo>> GetRows(List<ExcelColumnInfo> props)
        {
            while (_reader.Read())
            {
                yield return GetRowValues(props);
            }
        }

        private IEnumerable<CellWriteInfo> GetRowValues(List<ExcelColumnInfo> props)
        {
            for (int i = 0, column = 1; i < _reader.FieldCount; i++, column++)
            {
                if (_configuration.DynamicColumnFirst)
                {
                    var columnIndex = _reader.GetOrdinal(props[i].Key.ToString());
                    yield return new CellWriteInfo(_reader.GetValue(columnIndex), column, props[i]);
                }
                else
                {
                    yield return new CellWriteInfo(_reader.GetValue(i), column, props[i]);
                }
            }
        }
    }
}


