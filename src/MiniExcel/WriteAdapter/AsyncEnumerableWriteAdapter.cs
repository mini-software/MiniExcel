using MiniExcelLibs.Utils;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

#if NETSTANDARD2_0_OR_GREATER || NET
namespace MiniExcelLibs.WriteAdapter
{
    internal class AsyncEnumerableWriteAdapter<T> : IAsyncMiniExcelWriteAdapter
    {
        private readonly IAsyncEnumerable<T> _values;
        private readonly Configuration _configuration;
        private IAsyncEnumerator<T> _enumerator;
        private bool _empty;

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
                return new List<ExcelColumnInfo>();
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
                _enumerator = _values.GetAsyncEnumerator();
                if (!await _enumerator.MoveNextAsync())
                {
                    yield break;
                }
            }

            do
            {
                cancellationToken.ThrowIfCancellationRequested();
                yield return GetRowValuesAsync(_enumerator.Current, props);

            } while (await _enumerator.MoveNextAsync());
        }

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        public async static IAsyncEnumerable<CellWriteInfo> GetRowValuesAsync(T currentValue, List<ExcelColumnInfo> props)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            var column = 1;
            foreach (var prop in props)
            {
                if (prop == null)
                {
                    column++;
                    continue;
                }

                switch (currentValue)
                {
                    case IDictionary<string, object> genericDictionary:
                        yield return new CellWriteInfo(genericDictionary[prop.Key.ToString()], column, prop);
                        break;
                    case IDictionary dictionary:
                        yield return new CellWriteInfo(dictionary[prop.Key], column, prop);
                        break;
                    default:
                        yield return new CellWriteInfo(prop.Property.GetValue(currentValue), column, prop);
                        break;
                }

                column++;
            }
        }
    }
}
#endif