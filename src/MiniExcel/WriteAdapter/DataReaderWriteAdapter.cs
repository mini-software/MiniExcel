using MiniExcelLibs.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;

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

        public bool TryGetKnownCount(out int count)
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
                if (!_configuration.DynamicColumnFirst || 
                    _configuration.DynamicColumns.Any(d => string.Equals(d.Key, columnName, StringComparison.OrdinalIgnoreCase)))
                {
                    var prop = CustomPropertyHelper.GetColumnInfosFromDynamicConfiguration(columnName, _configuration);
                    props.Add(prop);
                }
            }
            return props;
        }

        public IEnumerable<IEnumerable<CellWriteInfo>> GetRows(List<ExcelColumnInfo> props, CancellationToken cancellationToken = default)
        {
            while (_reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                yield return GetRowValues(props);
            }
        }

        private IEnumerable<CellWriteInfo> GetRowValues(List<ExcelColumnInfo> props)
        {
            var column = 1;
            for (int i = 0; i < _reader.FieldCount; i++)
            {
                var prop = props[i];
                if (prop != null && !prop.ExcelIgnore)
                {
                    var columnIndex = _configuration.DynamicColumnFirst 
                        ? _reader.GetOrdinal(prop.Key.ToString()) : i;
                    
                    yield return new CellWriteInfo(_reader.GetValue(columnIndex), column, prop);
                    column++;
                }
            }
        }
    }
}


