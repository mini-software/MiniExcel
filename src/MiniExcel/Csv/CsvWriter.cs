using MiniExcelLibs.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs.Csv
{
    internal class CsvWriter : IExcelWriter, IDisposable
    {
        private readonly Stream _stream;
        private readonly CsvConfiguration _configuration;
        private readonly bool _printHeader;
        private object _value;
        private readonly StreamWriter _writer;
        private bool disposedValue;

        public CsvWriter(Stream stream, object value, IConfiguration configuration, bool printHeader)
        {
            this._stream = stream;
            this._configuration = configuration == null ? CsvConfiguration.DefaultConfiguration : (CsvConfiguration)configuration;
            this._printHeader = printHeader;
            this._value = value;
            this._writer = _configuration.StreamWriterFunc(_stream);
        }

        public void SaveAs()
        {
            var seperator = _configuration.Seperator.ToString();
            var newLine = _configuration.NewLine;
            {
                if (_value == null)
                {
                    _writer.Write("");
                    this._writer.Flush();
                    return;
                }

                var type = _value.GetType();

                if (_value is IDataReader dataReader)
                {
                    GenerateSheetByIDataReader(dataReader, seperator, newLine, _writer);
                }
                else if (_value is IEnumerable enumerable)
                {
                    GenerateSheetByIEnumerable(enumerable, seperator, newLine, _writer);
                }
                else if (_value is DataTable dataTable)
                {
                    GenerateSheetByDataTable(_writer, dataTable, seperator, newLine);
                }
                else
                {
                    throw new NotImplementedException($"Type {type?.Name} not Implemented. please issue for me.");
                }

                this._writer.Flush();
            }
        }

        private void GenerateSheetByIEnumerable(IEnumerable values, string seperator, string newLine, StreamWriter writer)
        {
            Type genericType = null;
            List<ExcelColumnInfo> props = null;
            string mode = null;

            var enumerator = values.GetEnumerator();
            var empty = !enumerator.MoveNext();
            if (empty)
            {
                // only when empty IEnumerable need to check this issue #133  https://github.com/shps951023/MiniExcel/issues/133
                genericType = TypeHelper.GetGenericIEnumerables(values).FirstOrDefault();
                if (genericType == null || genericType == typeof(object) // sometime generic type will be object, e.g: https://user-images.githubusercontent.com/12729184/132812859-52984314-44d1-4ee8-9487-2d1da159f1f0.png
                    || typeof(IDictionary<string, object>).IsAssignableFrom(genericType)
                    || typeof(IDictionary).IsAssignableFrom(genericType)
                    || typeof(KeyValuePair<string, object>).IsAssignableFrom(genericType))
                {
                    _writer.Write(newLine);
                    this._writer.Flush();
                    return;
                }

                mode = "Properties";
                props = CustomPropertyHelper.GetSaveAsProperties(genericType, _configuration);
            }
            else
            {
                var firstItem = enumerator.Current;
                if (firstItem is IDictionary<string, object> genericDic)
                {
                    mode = "IDictionary<string, object>";
                    props = CustomPropertyHelper.GetDictionaryColumnInfo(genericDic, null, _configuration);
                }
                else if (firstItem is IDictionary dic)
                {
                    mode = "IDictionary";
                    props = CustomPropertyHelper.GetDictionaryColumnInfo(null, dic, _configuration);
                    mode = "IDictionary";
                }
                else
                {
                    mode = "Properties";
                    genericType = firstItem.GetType();
                    props = CustomPropertyHelper.GetSaveAsProperties(genericType, _configuration);
                }
            }

            if (this._printHeader)
            {
                _writer.Write(string.Join(seperator, props.Select(s => CsvHelpers.ConvertToCsvValue(s?.ExcelColumnName, _configuration.AlwaysQuote, _configuration.Seperator))));
                _writer.Write(newLine);
            }

            if (!empty)
            {
                if (mode == "IDictionary<string, object>") //Dapper Row
                    GenerateSheetByDapperRow(_writer, enumerator, props.Select(x => x.Key.ToString()).ToList(), seperator, newLine);
                else if (mode == "IDictionary") //IDictionary
                    GenerateSheetByIDictionary(_writer, enumerator, props.Select(x => x.Key).ToList(), seperator, newLine);
                else if (mode == "Properties")
                    GenerateSheetByProperties(_writer, enumerator, props, seperator, newLine);
                else
                    throw new NotImplementedException($"Mode for genericType {genericType?.Name} not Implemented. please issue for me.");
            }
        }

        public void Insert()
        {
            SaveAs();
        }

        public async Task SaveAsAsync(CancellationToken cancellationToken = default(CancellationToken))
        {
            await Task.Run(() => SaveAs(), cancellationToken).ConfigureAwait(false);
        }

        private void GenerateSheetByIDataReader(IDataReader reader, string seperator, string newLine, StreamWriter writer)
        {
            int fieldCount = reader.FieldCount;
            if (fieldCount == 0)
                throw new InvalidDataException("fieldCount is 0");

            if (this._printHeader)
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    var columnName = reader.GetName(i);

                    if (i != 0)
                        writer.Write(seperator);
                    writer.Write(CsvHelpers.ConvertToCsvValue(ToCsvString(columnName, null), _configuration.AlwaysQuote, _configuration.Seperator));
                }
                writer.Write(newLine);
            }

            while (reader.Read())
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    var cellValue = reader.GetValue(i);
                    if (i != 0)
                        writer.Write(seperator);
                    writer.Write(CsvHelpers.ConvertToCsvValue(ToCsvString(cellValue, null), _configuration.AlwaysQuote, _configuration.Seperator));
                }
                writer.Write(newLine);
            }
        }

        private void GenerateSheetByDataTable(StreamWriter writer, DataTable dt, string seperator, string newLine)
        {
            if (_printHeader)
            {
                writer.Write(string.Join(seperator, dt.Columns.Cast<DataColumn>().Select(s => CsvHelpers.ConvertToCsvValue(s.Caption ?? s.ColumnName, _configuration.AlwaysQuote, _configuration.Seperator))));
                writer.Write(newLine);
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var first = true;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    var cellValue = CsvHelpers.ConvertToCsvValue(ToCsvString(dt.Rows[i][j], null), _configuration.AlwaysQuote, _configuration.Seperator);
                    if (!first)
                        writer.Write(seperator);
                    writer.Write(cellValue);
                    first = false;
                }
                writer.Write(newLine);
            }
        }

        private void GenerateSheetByProperties(StreamWriter writer, IEnumerator value, List<ExcelColumnInfo> props, string seperator, string newLine)
        {
            do
            {
                var v = value.Current;
                var values = props.Select(s => CsvHelpers.ConvertToCsvValue(ToCsvString(s?.Property.GetValue(v), s), _configuration.AlwaysQuote, _configuration.Seperator));
                writer.Write(string.Join(seperator, values));
                writer.Write(newLine);
            } while (value.MoveNext());
        }

        private void GenerateSheetByIDictionary(StreamWriter writer, IEnumerator value, List<object> keys, string seperator, string newLine)
        {
            do
            {
                var v = (IDictionary)value.Current;
                var values = keys.Select(key => CsvHelpers.ConvertToCsvValue(ToCsvString(v[key], null), _configuration.AlwaysQuote, _configuration.Seperator));
                writer.Write(string.Join(seperator, values));
                writer.Write(newLine);
            } while (value.MoveNext());
        }

        private void GenerateSheetByDapperRow(StreamWriter writer, IEnumerator value, List<string> keys, string seperator, string newLine)
        {
            do
            {
                var v = (IDictionary<string, object>)value.Current;
                var values = keys.Select(key => CsvHelpers.ConvertToCsvValue(ToCsvString(v[key], null), _configuration.AlwaysQuote, _configuration.Seperator));
                writer.Write(string.Join(seperator, values));
                writer.Write(newLine);
            } while (value.MoveNext());
        }

        public string ToCsvString(object value, ExcelColumnInfo p)
        {
            if (value == null)
                return "";

            if (value is DateTime dateTime)
            {
                if (p?.ExcelFormat != null)
                {
                    return dateTime.ToString(p.ExcelFormat, _configuration.Culture);
                }
                return _configuration.Culture.Equals(CultureInfo.InvariantCulture) ? dateTime.ToString("yyyy-MM-dd HH:mm:ss", _configuration.Culture) : dateTime.ToString(_configuration.Culture);
            }
            if (p?.ExcelFormat != null && value is IFormattable formattableValue)
            {
                return formattableValue.ToString(p.ExcelFormat, _configuration.Culture);
            }

            return Convert.ToString(value, _configuration.Culture);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    this._writer.Dispose();
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        ~CsvWriter()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: false);
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}