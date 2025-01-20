using MiniExcelLibs.Utils;
using MiniExcelLibs.WriteAdapter;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
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
            if (_value == null)
            {
                _writer.Write("");
                _writer.Flush();
                return;
            }

            WriteValues(_writer, _value);
            _writer.Flush();
        }

        public void Insert(bool overwriteSheet = false)
        {
            SaveAs();
        }

        private void AppendColumn(StringBuilder rowBuilder, CellWriteInfo column)
        {
            rowBuilder.Append(CsvHelpers.ConvertToCsvValue(ToCsvString(column.Value, column.Prop), _configuration.AlwaysQuote, _configuration.Seperator));
            rowBuilder.Append(_configuration.Seperator);
        }

        private void RemoveTrailingSeparator(StringBuilder rowBuilder)
        {
            if (rowBuilder.Length == 0)
            {
                return;
            }
            rowBuilder.Remove(rowBuilder.Length - 1, 1);
        }

        private string GetHeader(List<ExcelColumnInfo> props) => string.Join(
            _configuration.Seperator.ToString(),
            props.Select(s => CsvHelpers.ConvertToCsvValue(s?.ExcelColumnName, _configuration.AlwaysQuote, _configuration.Seperator)));

        private void WriteValues(StreamWriter writer, object values)
        {
            IMiniExcelWriteAdapter writeAdapter = MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);

            var props = writeAdapter.GetColumns();
            if (props == null)
            {
                _writer.Write(_configuration.NewLine);
                _writer.Flush();
                return;
            }

            if (_printHeader)
            {
                _writer.Write(GetHeader(props));
                _writer.Write(_configuration.NewLine);
            }

            var rowBuilder = new StringBuilder();
            if (writeAdapter != null)
            {
                foreach (var row in writeAdapter.GetRows(props))
                {
                    rowBuilder.Clear();
                    foreach (var column in row)
                    {
                        AppendColumn(rowBuilder, column);
                    }
                    RemoveTrailingSeparator(rowBuilder);
                    _writer.Write(rowBuilder.ToString());
                    _writer.Write(_configuration.NewLine);
                }
            }
        }

        private async Task WriteValuesAsync(StreamWriter writer, object values, string seperator, string newLine, CancellationToken cancellationToken)
        {
#if NETSTANDARD2_0_OR_GREATER || NET
            IMiniExcelWriteAdapter writeAdapter = null;
            if (!MiniExcelWriteAdapterFactory.TryGetAsyncWriteAdapter(values, _configuration, out var asyncWriteAdapter))
            {
                writeAdapter = MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);
            }
            var props = writeAdapter?.GetColumns() ?? await asyncWriteAdapter.GetColumnsAsync();
#else
            IMiniExcelWriteAdapter writeAdapter =  MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);
            var props = writeAdapter.GetColumns();
#endif
            if (props == null)
            {
                await _writer.WriteAsync(_configuration.NewLine);
                await _writer.FlushAsync();
                return;
            }
            if (_printHeader)
            {
                await _writer.WriteAsync(GetHeader(props));
                await _writer.WriteAsync(newLine);
            }
            var rowBuilder = new StringBuilder();
            if (writeAdapter != null)
            {
                foreach (var row in writeAdapter.GetRows(props, cancellationToken))
                {
                    rowBuilder.Clear();
                    foreach (var column in row)
                    {
                        AppendColumn(rowBuilder, column);
                    }
                    RemoveTrailingSeparator(rowBuilder);
                    await _writer.WriteAsync(rowBuilder.ToString());
                    await _writer.WriteAsync(newLine);
                }
            }
#if NETSTANDARD2_0_OR_GREATER || NET
            else
            {
                await foreach (var row in asyncWriteAdapter.GetRowsAsync(props, cancellationToken))
                {
                    rowBuilder.Clear();
                    await foreach (var column in row)
                    {
                        AppendColumn(rowBuilder, column);
                    }
                    RemoveTrailingSeparator(rowBuilder);
                    await _writer.WriteAsync(rowBuilder.ToString());
                    await _writer.WriteAsync(newLine);
                }
            }
#endif
        }

        public async Task SaveAsAsync(CancellationToken cancellationToken = default)
        {
            var seperator = _configuration.Seperator.ToString();
            var newLine = _configuration.NewLine;

            if (_value == null)
            {
                await _writer.WriteAsync("");
                await _writer.FlushAsync();
                return;
            }

            await WriteValuesAsync(_writer, _value, seperator, newLine, cancellationToken);
            await _writer.FlushAsync();
        }

        public async Task InsertAsync(bool overwriteSheet = false, CancellationToken cancellationToken = default)
        {
            await SaveAsAsync(cancellationToken);
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