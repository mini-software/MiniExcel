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
        private readonly StreamWriter _writer;
        private readonly CsvConfiguration _configuration;
        private readonly bool _printHeader;
        private readonly object _value;
        private bool _disposedValue;

        public CsvWriter(Stream stream, object value, IConfiguration configuration, bool printHeader)
        {
            _configuration = configuration == null ? CsvConfiguration.DefaultConfiguration : (CsvConfiguration)configuration;
            _printHeader = printHeader;
            _value = value;
            _writer = _configuration.StreamWriterFunc(stream);
        }

        public int[] SaveAs()
        {
            if (_value == null)
            {
                _writer.Write("");
                _writer.Flush();
                return new int[0];
            }

            var rowsWritten = WriteValues(_value);
            _writer.Flush();

            return new[] { rowsWritten };
        }

        public int Insert(bool overwriteSheet = false)
        {
            return SaveAs().FirstOrDefault();
        }

        private void AppendColumn(StringBuilder rowBuilder, CellWriteInfo column)
        {
            rowBuilder.Append(CsvHelpers.ConvertToCsvValue(ToCsvString(column.Value, column.Prop), _configuration));
            rowBuilder.Append(_configuration.Seperator);
        }

        private static void RemoveTrailingSeparator(StringBuilder rowBuilder)
        {
            if (rowBuilder.Length == 0)
                return;
            
            rowBuilder.Remove(rowBuilder.Length - 1, 1);
        }

        private string GetHeader(List<ExcelColumnInfo> props) => string.Join(
            _configuration.Seperator.ToString(),
            props.Select(s => CsvHelpers.ConvertToCsvValue(s?.ExcelColumnName, _configuration)));

        private int WriteValues(object values)
        {
            var writeAdapter = MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);

            var props = writeAdapter.GetColumns();
            if (props == null)
            {
                _writer.Write(_configuration.NewLine);
                _writer.Flush();
                return 0;
            }

            if (_printHeader)
            {
                _writer.Write(GetHeader(props));
                _writer.Write(_configuration.NewLine);
            }

            if (writeAdapter == null) 
                return 0;
            
            var rowBuilder = new StringBuilder();
            var rowsWritten = 0;

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
                    
                rowsWritten++;
            }
            return rowsWritten;
        }

        private async Task<int> WriteValuesAsync(StreamWriter writer, object values, string seperator, string newLine, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            
#if NETSTANDARD2_0_OR_GREATER || NET
            IMiniExcelWriteAdapter writeAdapter = null;
            if (!MiniExcelWriteAdapterFactory.TryGetAsyncWriteAdapter(values, _configuration, out var asyncWriteAdapter))
            {
                writeAdapter = MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);
            }
            var props = writeAdapter != null ? writeAdapter.GetColumns() : await asyncWriteAdapter.GetColumnsAsync().ConfigureAwait(false);
#else
            IMiniExcelWriteAdapter writeAdapter =  MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);
            var props = writeAdapter.GetColumns();
#endif
            if (props == null)
            {
                await _writer.WriteAsync(_configuration.NewLine
#if NET5_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
                    ).ConfigureAwait(false);
                await _writer.FlushAsync(
#if NET8_0_OR_GREATER
                    cancellationToken
#endif
                    ).ConfigureAwait(false);
                return 0;
            }
            
            if (_printHeader)
            {
                await _writer.WriteAsync(GetHeader(props)
#if NET5_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
                    ).ConfigureAwait(false);
                await _writer.WriteAsync(newLine
#if NET5_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
                    ).ConfigureAwait(false);
            }
            
            var rowBuilder = new StringBuilder();
            var rowsWritten = 0;
            
            if (writeAdapter != null)
            {
                foreach (var row in writeAdapter.GetRows(props, cancellationToken))
                {
                    rowBuilder.Clear();
                    foreach (var column in row)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        AppendColumn(rowBuilder, column);
                    }
                    
                    RemoveTrailingSeparator(rowBuilder);
                    await _writer.WriteAsync(rowBuilder.ToString()
#if NET5_0_OR_GREATER
                        .AsMemory(), cancellationToken
#endif
                    ).ConfigureAwait(false);
                    await _writer.WriteAsync(newLine
#if NET5_0_OR_GREATER
                        .AsMemory(), cancellationToken
#endif
                    ).ConfigureAwait(false);
                    
                    rowsWritten++;
                }
            }
#if NETSTANDARD2_0_OR_GREATER || NET
            else
            {
                await foreach (var row in asyncWriteAdapter.GetRowsAsync(props, cancellationToken).ConfigureAwait(false))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    rowBuilder.Clear();
                    
                    await foreach (var column in row.ConfigureAwait(false))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        AppendColumn(rowBuilder, column);
                    }
                    
                    RemoveTrailingSeparator(rowBuilder);
                    await _writer.WriteAsync(rowBuilder.ToString()
#if NET5_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
                    ).ConfigureAwait(false);
                    await _writer.WriteAsync(newLine
#if NET5_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
                    ).ConfigureAwait(false);
                    
                    rowsWritten++;
                }
            }
#endif
            return rowsWritten;
        }

        public async Task<int[]> SaveAsAsync(CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            
            var seperator = _configuration.Seperator.ToString();
            var newLine = _configuration.NewLine;

            if (_value == null)
            {
                await _writer.WriteAsync(""
#if NET5_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
                    ).ConfigureAwait(false);
                await _writer.FlushAsync(
#if NET8_0_OR_GREATER
                    cancellationToken
#endif
                    ).ConfigureAwait(false);
                return new int[0];
            }

            var rowsWritten = await WriteValuesAsync(_writer, _value, seperator, newLine, cancellationToken).ConfigureAwait(false);
            await _writer.FlushAsync(
#if NET8_0_OR_GREATER
                    cancellationToken
#endif
                ).ConfigureAwait(false);
         
            return new[] { rowsWritten };
        }

        public async Task<int> InsertAsync(bool overwriteSheet = false, CancellationToken cancellationToken = default)
        {
            var rowsWritten = await SaveAsAsync(cancellationToken).ConfigureAwait(false);
            return rowsWritten.FirstOrDefault();
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
                return _configuration.Culture.Equals(CultureInfo.InvariantCulture) 
                    ? dateTime.ToString("yyyy-MM-dd HH:mm:ss", _configuration.Culture) 
                    : dateTime.ToString(_configuration.Culture);
            }
            
            if (p?.ExcelFormat != null && value is IFormattable formattableValue)
                return formattableValue.ToString(p.ExcelFormat, _configuration.Culture);

            return Convert.ToString(value, _configuration.Culture);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    _writer.Dispose();
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                _disposedValue = true;
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