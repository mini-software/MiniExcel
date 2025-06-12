using MiniExcelLibs.Exceptions;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;

namespace MiniExcelLibs.Csv
{
    internal partial class CsvReader : IExcelReader
    {
        private Stream _stream;
        private CsvConfiguration _config;

        public CsvReader(Stream stream, IConfiguration configuration)
        {
            _stream = stream;
            _config = configuration == null ? CsvConfiguration.DefaultConfiguration : (CsvConfiguration)configuration;
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public async IAsyncEnumerable<IDictionary<string, object>> QueryAsync(bool useHeaderRow, string sheetName, string startCell, [EnumeratorCancellation] CancellationToken ct = default)
        {
            if (startCell != "A1")
                throw new NotImplementedException("CSV does not implement parameter startCell");

            if (_stream.CanSeek)
                _stream.Position = 0;

            var reader = _config.StreamReaderFunc(_stream);
            var firstRow = true;
            var headRows = new Dictionary<int, string>();

            string row;
            for (var rowIndex = 1; (row = await reader.ReadLineAsync(
#if NET7_0_OR_GREATER
ct
#endif
                ).ConfigureAwait(false)) != null; rowIndex++)
            {
                string finalRow = row;
                if (_config.ReadLineBreaksWithinQuotes)
                {
                    while (finalRow.Count(c => c == '"') % 2 != 0)
                    {
                        var nextPart = await reader.ReadLineAsync(
#if NET7_0_OR_GREATER
ct
#endif
                            ).ConfigureAwait(false);
                        if (nextPart == null)
                        {
                            break;
                        }
                        finalRow = string.Concat(finalRow, _config.NewLine, nextPart);
                    }
                }
                var read = Split(finalRow);

                // invalid row check
                if (read.Length < headRows.Count)
                {
                    var colIndex = read.Length;
                    var headers = headRows.ToDictionary(x => x.Value, x => x.Key);
                    var rowValues = read
                        .Select((x, i) => new KeyValuePair<string, object>(headRows[i], x))
                        .ToDictionary(x => x.Key, x => x.Value);

                    throw new ExcelColumnNotFoundException(columnIndex: null, headRows[colIndex], null, rowIndex, headers, rowValues, $"Csv read error: Column {colIndex} not found in Row {rowIndex}");
                }

                //header
                if (useHeaderRow)
                {
                    if (firstRow)
                    {
                        firstRow = false;
                        for (int i = 0; i <= read.Length - 1; i++)
                            headRows.Add(i, read[i]);
                        continue;
                    }

                    var headCell = CustomPropertyHelper.GetEmptyExpandoObject(headRows);
                    for (int i = 0; i <= read.Length - 1; i++)
                        headCell[headRows[i]] = read[i];

                    yield return headCell;
                    continue;
                }

                //body
                if (firstRow) // record first row as reference
                {
                    firstRow = false;
                    for (int i = 0; i <= read.Length - 1; i++)
                        headRows.Add(i, $"c{i + 1}");
                }

                var cell = CustomPropertyHelper.GetEmptyExpandoObject(read.Length - 1, 0);
                if (_config.ReadEmptyStringAsNull)
                {
                    for (int i = 0; i <= read.Length - 1; i++)
                        cell[ColumnHelper.GetAlphabetColumnName(i)] = read[i]?.Length == 0 ? null : read[i];
                }
                else
                {
                    for (int i = 0; i <= read.Length - 1; i++)
                        cell[ColumnHelper.GetAlphabetColumnName(i)] = read[i];
                }

                yield return cell;
            }
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public IAsyncEnumerable<T> QueryAsync<T>(string sheetName, string startCell, bool hasHeader, CancellationToken ct = default) where T : class, new()
        {
            var dynamicRecords = QueryAsync(false, sheetName, startCell, ct);
            return ExcelOpenXmlSheetReader.QueryImplAsync<T>(dynamicRecords, startCell, hasHeader, _config, ct);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public IAsyncEnumerable<IDictionary<string, object>> QueryRangeAsync(bool useHeaderRow, string sheetName, string startCell, string endCell, CancellationToken ct = default)
        {
            throw new NotImplementedException("CSV does not implement QueryRange");
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public IAsyncEnumerable<T> QueryRangeAsync<T>(string sheetName, string startCell, string endCell, bool hasHeader, CancellationToken ct = default) where T : class, new()
        {
            var dynamicRecords = QueryRangeAsync(false, sheetName, startCell, endCell, ct);
            return ExcelOpenXmlSheetReader.QueryImplAsync<T>(dynamicRecords, startCell, hasHeader, this._config, ct);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public IAsyncEnumerable<IDictionary<string, object>> QueryRangeAsync(bool useHeaderRow, string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, CancellationToken ct = default)
        {
            throw new NotImplementedException("CSV does not implement QueryRange");
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public IAsyncEnumerable<T> QueryRangeAsync<T>(string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, bool hasHeader, CancellationToken ct = default) where T : class, new()
        {
            var dynamicRecords = QueryRangeAsync(false, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, ct);
            return ExcelOpenXmlSheetReader.QueryImplAsync<T>(dynamicRecords, ReferenceHelper.ConvertXyToCell(startRowIndex, startColumnIndex), hasHeader, this._config, ct);
        }

        private string[] Split(string row)
        {
            if (_config.SplitFn != null)
            {
                return _config.SplitFn(row);
            }
            else
            {
                //this code from S.O : https://stackoverflow.com/a/11365961/9131476
                return Regex.Split(row, $"[\t{_config.Seperator}](?=(?:[^\"]|\"[^\"]*\")*$)")
                    .Select(s => Regex.Replace(s.Replace("\"\"", "\""), "^\"|\"$", ""))
                    .ToArray();
            }
        }

        public void Dispose()
        {
        }
    }
}
