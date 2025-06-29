using System.Runtime.CompilerServices;
using MiniExcelLib.Core.Abstractions;
using MiniExcelLib.Core.Reflection;

namespace MiniExcelLib.Core.WriteAdapters;

internal class MiniExcelDataReaderWriteAdapter(IMiniExcelDataReader reader, MiniExcelBaseConfiguration configuration) : IMiniExcelWriteAdapterAsync
{
    private readonly IMiniExcelDataReader _reader = reader;
    private readonly MiniExcelBaseConfiguration _configuration = configuration;

    public async Task<List<MiniExcelColumnInfo>?> GetColumnsAsync()
    {
        List<MiniExcelColumnInfo> props = [];
        for (var i = 0; i < _reader.FieldCount; i++)
        {
            var columnName = await _reader.GetNameAsync(i).ConfigureAwait(false);

            if (!_configuration.DynamicColumnFirst)
            {
                var prop = CustomPropertyHelper.GetColumnInfosFromDynamicConfiguration(columnName, _configuration);
                props.Add(prop);
                continue;
            }

            if (_configuration.DynamicColumns?.Any(a => string.Equals(a.Key, columnName, StringComparison.OrdinalIgnoreCase)) ?? false)
            {
                var prop = CustomPropertyHelper.GetColumnInfosFromDynamicConfiguration(columnName, _configuration);
                props.Add(prop);
            }
        }
        return props;
    }

    public async IAsyncEnumerable<IAsyncEnumerable<CellWriteInfo>> GetRowsAsync(List<MiniExcelColumnInfo> props, [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        while (await _reader.ReadAsync(cancellationToken).ConfigureAwait(false))
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return GetRowValuesAsync(props);
        }
    }

    private async IAsyncEnumerable<CellWriteInfo> GetRowValuesAsync(List<MiniExcelColumnInfo> props)
    {
        for (int i = 0, column = 1; i < _reader.FieldCount; i++, column++)
        {
            if (_configuration.DynamicColumnFirst)
            {
                var columnIndex = _reader.GetOrdinal(props[i].Key.ToString());
                yield return new CellWriteInfo(await _reader.GetValueAsync(columnIndex).ConfigureAwait(false), column, props[i]);
            }
            else
            {
                yield return new CellWriteInfo(await _reader.GetValueAsync(i).ConfigureAwait(false), column, props[i]);
            }
        }
    }
}