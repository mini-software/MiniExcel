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

namespace MiniExcelLibs.Csv
{
    internal class CsvReader : IExcelReader
    {
        private Stream _stream;
        private CsvConfiguration _config;

        public CsvReader(Stream stream, IConfiguration configuration)
        {
            _stream = stream;
            _config = configuration == null ? CsvConfiguration.DefaultConfiguration : (CsvConfiguration)configuration;
        }

        public IEnumerable<IDictionary<string, object>> Query(bool useHeaderRow, string sheetName, string startCell)
        {
            if (startCell != "A1")
                throw new NotImplementedException("CSV does not implement parameter startCell");

            if (_stream.CanSeek)
                _stream.Position = 0;

            var reader = _config.StreamReaderFunc(_stream);
            var firstRow = true;
            var headRows = new Dictionary<int, string>();

            string row;
            for (var rowIndex = 1; (row = reader.ReadLine()) != null; rowIndex++)
            {
                if (string.IsNullOrWhiteSpace(row))
                    continue;
                
                string finalRow = row;
                if (_config.ReadLineBreaksWithinQuotes)
                {
                    while (finalRow.Count(c => c == '"') % 2 != 0)
                    {
                        var nextPart = reader.ReadLine();
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

        public IEnumerable<T> Query<T>(string sheetName, string startCell, bool hasHeader) where T : class, new()
        {
            var dynamicRecords = Query(false, sheetName, startCell);
            return ExcelOpenXmlSheetReader.QueryImpl<T>(dynamicRecords, startCell, hasHeader, _config);
        }

        public IEnumerable<IDictionary<string, object>> QueryRange(bool useHeaderRow, string sheetName, string startCell, string endCell)
        {
            throw new NotImplementedException("CSV does not implement QueryRange");
        }

        public IEnumerable<T> QueryRange<T>(string sheetName, string startCell, string endCell, bool hasHeader) where T : class, new()
        {
            var dynamicRecords = QueryRange(false, sheetName, startCell, endCell);
            return ExcelOpenXmlSheetReader.QueryImpl<T>(dynamicRecords, startCell, hasHeader, this._config);
        }

        public IEnumerable<IDictionary<string, object>> QueryRange(bool useHeaderRow, string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex)
        {
            throw new NotImplementedException("CSV does not implement QueryRange");
        }

        public IEnumerable<T> QueryRange<T>(string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, bool hasHeader) where T : class, new()
        {
            var dynamicRecords = QueryRange(false, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
            return ExcelOpenXmlSheetReader.QueryImpl<T>(dynamicRecords, ReferenceHelper.ConvertXyToCell(startRowIndex, startColumnIndex), hasHeader, this._config);
        }

        public Task<IEnumerable<IDictionary<string, object>>> QueryAsync(bool useHeaderRow, string sheetName, string startCell, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => Query(useHeaderRow, sheetName, startCell), cancellationToken);
        }

        public async Task<IEnumerable<T>> QueryAsync<T>(string sheetName, string startCell, bool hasHeader, CancellationToken cancellationToken = default) where T : class, new()
        {
            return await Task.Run(() => Query<T>(sheetName, startCell, hasHeader), cancellationToken).ConfigureAwait(false);
        }

        public Task<IEnumerable<IDictionary<string, object>>> QueryRangeAsync(bool useHeaderRow, string sheetName, string startCell, string endCel, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => QueryRange(useHeaderRow, sheetName, startCell, endCel), cancellationToken);
        }

        public async Task<IEnumerable<T>> QueryRangeAsync<T>(string sheetName, string startCell, string endCel, bool hasHeader, CancellationToken cancellationToken = default) where T : class, new()
        {
            return await Task.Run(() => QueryRange<T>(sheetName, startCell, endCel, hasHeader), cancellationToken).ConfigureAwait(false);
        }

        public Task<IEnumerable<IDictionary<string, object>>> QueryRangeAsync(bool useHeaderRow, string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => QueryRange(useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex), cancellationToken);
        }

        public async Task<IEnumerable<T>> QueryRangeAsync<T>(string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, bool hasHeader, CancellationToken cancellationToken = default) where T : class, new()
        {
            return await Task.Run(() => QueryRange<T>(sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, hasHeader), cancellationToken).ConfigureAwait(false);
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
