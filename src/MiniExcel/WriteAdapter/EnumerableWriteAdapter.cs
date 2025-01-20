using MiniExcelLibs.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace MiniExcelLibs.WriteAdapter
{
    internal class EnumerableWriteAdapter : IMiniExcelWriteAdapter
    {
        private readonly IEnumerable _values;
        private readonly Configuration _configuration;
        private readonly Type _genericType;

        private IEnumerator _enumerator;
        private bool _empty;

        public EnumerableWriteAdapter(IEnumerable values, Configuration configuration)
        {
            _values = values;
            _configuration = configuration;
            _genericType = TypeHelper.GetGenericIEnumerables(values).FirstOrDefault();
        }

        public bool TryGetKnownCount(out int count)
        {
            count = 0;
            if (_values is ICollection collection)
            {
                count = collection.Count;
                return true;
            }

            return false;
        }

        public List<ExcelColumnInfo> GetColumns()
        {
            if (CustomPropertyHelper.TryGetTypeColumnInfo(_genericType, _configuration, out var props))
            {
                return props;
            }

            _enumerator = _values.GetEnumerator();
            if (!_enumerator.MoveNext())
            {
                _empty = true;
                return null;
            }
            return CustomPropertyHelper.GetColumnInfoFromValue(_enumerator.Current, _configuration);           
        }

        public IEnumerable<IEnumerable<CellWriteInfo>> GetRows(List<ExcelColumnInfo> props, CancellationToken cancellationToken = default)
        {
            if (_empty)
            {
                yield break;
            }

            if (_enumerator is null)
            {
                _enumerator = _values.GetEnumerator();
                if (!_enumerator.MoveNext())
                {
                    yield break;
                }
            }

            do
            {
                cancellationToken.ThrowIfCancellationRequested();
                yield return GetRowValues(_enumerator.Current, props);
            } while (_enumerator.MoveNext());
        }


        public static IEnumerable<CellWriteInfo> GetRowValues(object currentValue, List<ExcelColumnInfo> props)
        {
            var column = 1;
            foreach (var prop in props)
            {
                object cellValue;
                if (prop == null)
                {
                    cellValue = null;
                }
                else if (currentValue is IDictionary<string, object> genericDictionary)
                {
                    cellValue = genericDictionary[prop.Key.ToString()];
                }
                else if (currentValue is IDictionary dictionary)
                {
                    cellValue = dictionary[prop.Key];
                }
                else 
                {
                    cellValue = prop.Property.GetValue(currentValue);
                }
                yield return new CellWriteInfo(cellValue, column, prop);
                column++;
            }
        }
    }
}
