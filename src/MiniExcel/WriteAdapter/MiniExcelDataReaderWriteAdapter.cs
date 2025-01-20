using MiniExcelLibs.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

#if NETSTANDARD2_0_OR_GREATER || NET
namespace MiniExcelLibs.WriteAdapter
{
    internal class MiniExcelDataReaderWriteAdapter : IAsyncMiniExcelWriteAdapter
    {
        private readonly IMiniExcelDataReader _reader;
        private readonly Configuration _configuration;

        public MiniExcelDataReaderWriteAdapter(IMiniExcelDataReader reader, Configuration configuration)
        {
            _reader = reader;
            _configuration = configuration;
        }

        public async Task<List<ExcelColumnInfo>> GetColumnsAsync()
        {
            var props = new List<ExcelColumnInfo>();
            for (var i = 0; i < _reader.FieldCount; i++)
            {
                var columnName = await _reader.GetNameAsync(i);

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

        public async IAsyncEnumerable<IAsyncEnumerable<CellWriteInfo>> GetRowsAsync(List<ExcelColumnInfo> props, [EnumeratorCancellation] CancellationToken cancellationToken)
        {
            while (await _reader.ReadAsync())
            {
                cancellationToken.ThrowIfCancellationRequested();
                yield return GetRowValuesAsync(props);
            }
        }

        private async IAsyncEnumerable<CellWriteInfo> GetRowValuesAsync(List<ExcelColumnInfo> props)
        {
            for (int i = 0, column = 1; i < _reader.FieldCount; i++, column++)
            {
                if (_configuration.DynamicColumnFirst)
                {
                    var columnIndex = _reader.GetOrdinal(props[i].Key.ToString());
                    yield return new CellWriteInfo(await _reader.GetValueAsync(columnIndex), column, props[i]);
                }
                else
                {
                    yield return new CellWriteInfo(await _reader.GetValueAsync(i), column, props[i]);
                }
            }
        }
    }
}
#endif
