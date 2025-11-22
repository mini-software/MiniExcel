using System;
using MiniExcelLibs.Utils;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

#if NETSTANDARD2_0_OR_GREATER || NET
namespace MiniExcelLibs.WriteAdapter
{
    internal class AsyncEnumerableWriteAdapter<T> : IAsyncMiniExcelWriteAdapter, IAsyncDisposable
    {
        private readonly IAsyncEnumerable<T> _values;
        private readonly Configuration _configuration;
        
        private IAsyncEnumerator<T> _enumerator;
        private bool _empty;
        private bool _disposed;

        public AsyncEnumerableWriteAdapter(IAsyncEnumerable<T> values, Configuration configuration)
        {
            _values = values;
            _configuration = configuration;
        }

        public async Task<List<ExcelColumnInfo>> GetColumnsAsync()
        {
            if (CustomPropertyHelper.TryGetTypeColumnInfo(typeof(T), _configuration, out var props))
            {
                return props;
            }

            _enumerator = _values.GetAsyncEnumerator();
            if (!await _enumerator.MoveNextAsync())
            {
                _empty = true;
                return null;
            }
            return CustomPropertyHelper.GetColumnInfoFromValue(_enumerator.Current, _configuration);
        }

        public async IAsyncEnumerable<IAsyncEnumerable<CellWriteInfo>> GetRowsAsync(List<ExcelColumnInfo> props, [EnumeratorCancellation] CancellationToken cancellationToken)
        {
            if (_empty)
            {
                yield break;
            }

            if (_enumerator is null)
            {
                _enumerator = _values.GetAsyncEnumerator(cancellationToken);
                if (!await _enumerator.MoveNextAsync())
                {
                    yield break;
                }
            }

            do
            {
                cancellationToken.ThrowIfCancellationRequested();
                yield return GetRowValuesAsync(_enumerator.Current, props);

            }
            while (await _enumerator.MoveNextAsync());
        }

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        private static async IAsyncEnumerable<CellWriteInfo> GetRowValuesAsync(T currentValue, List<ExcelColumnInfo> props)
#pragma warning restore CS1998
        {
            var column = 0;
            foreach (var prop in props)
            {
                column++;
            
                if (prop is null)
                    continue;

                yield return currentValue switch
                {
                    IDictionary<string, object> genericDictionary => new CellWriteInfo(genericDictionary[prop.Key.ToString()], column, prop),
                    IDictionary dictionary => new CellWriteInfo(dictionary[prop.Key], column, prop),
                    _ => new CellWriteInfo(prop.Property.GetValue(currentValue), column, prop)
                };
            }
        }

        public async ValueTask DisposeAsync()
        {
            if (!_disposed)
            {
                if (_enumerator != null)
                {
                    await _enumerator.DisposeAsync().ConfigureAwait(false);
                }
                _disposed = true;
            }
        }
    }
}
#endif