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
            this._stream = stream;
            this._config = configuration == null ? CsvConfiguration.DefaultConfiguration : (CsvConfiguration)configuration;
        }
        public IEnumerable<IDictionary<string, object>> Query(bool useHeaderRow, string sheetName, string startCell)
        {
            if (startCell != "A1")
                throw new NotImplementedException("CSV not Implement startCell");
            if (_stream.CanSeek)
                _stream.Position = 0;
            var reader = _config.StreamReaderFunc(_stream);
            {
                string[] read;
                var firstRow = true;
                Dictionary<int, string> headRows = new Dictionary<int, string>();
                string row;
                for (var rowIndex = 1; (row = reader.ReadLine()) != null; rowIndex++)
                {

                    string finalRow = row;
                    if (_config.ReadLineBreaksWithinQuotes)
                    {
                        while (finalRow.Count(c => c == '"') % 2 != 0)
                        {
                            var nextPart = reader.ReadLine();
                            if (nextPart == null ) {
                                break;
                            }
							finalRow = string.Concat( finalRow, _config.NewLine, nextPart );
                        }
                    }
                    read = Split(finalRow);

                    // invalid row check
                    if (read.Length < headRows.Count)
                    {
                        var colIndex = read.Length;
                        var headers = headRows.ToDictionary(x => x.Value, x => x.Key);
                        var rowValues = read.Select((x, i) => new { Key = headRows[i], Value = x }).ToDictionary(x => x.Key, x => (object)x.Value);
                        throw new Exceptions.ExcelColumnNotFoundException(null,
                            headRows[colIndex], null, rowIndex, headers, rowValues, $"Csv read error, Column: {colIndex} not found in Row: {rowIndex}");
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

                        var cell = CustomPropertyHelper.GetEmptyExpandoObject(headRows);
                        for (int i = 0; i <= read.Length - 1; i++)
                            cell[headRows[i]] = read[i];

                        yield return cell;
                        continue;
                    }


                    //body
                    {
                        // record first row as reference
                        if (firstRow)
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
            }
        }
        public IEnumerable<T> Query<T>(string sheetName, string startCell) where T : class, new()
        {
            return ExcelOpenXmlSheetReader.QueryImpl<T>(Query(false, sheetName, startCell), startCell, this._config);
        }

        private string[] Split(string row)
        {
            if (_config.SplitFn != null)
            {
                return _config.SplitFn(row);
            }
            else
            {
                return Regex.Split(row, $"[\t{_config.Seperator}](?=(?:[^\"]|\"[^\"]*\")*$)")
                    .Select(s => Regex.Replace(s.Replace("\"\"", "\""), "^\"|\"$", "")).ToArray();
                //this code from S.O : https://stackoverflow.com/a/11365961/9131476
            }
        }

        public Task<IEnumerable<IDictionary<string, object>>> QueryAsync(bool UseHeaderRow, string sheetName, string startCell, CancellationToken cancellationToken = default(CancellationToken))
        {
            return Task.Run(() => Query(UseHeaderRow, sheetName, startCell), cancellationToken);
        }

        public Task<IEnumerable<T>> QueryAsync<T>(string sheetName, string startCell, CancellationToken cancellationToken = default(CancellationToken)) where T : class, new()
        {
            return Task.Run(() => Query<T>(sheetName, startCell), cancellationToken);
        }

        public void Dispose()
        {
        }

        #region Range
        public IEnumerable<IDictionary<string, object>> QueryRange(bool useHeaderRow, string sheetName, string startCell, string endCell)
        {
            if (startCell != "A1")
                throw new NotImplementedException("CSV not Implement startCell");
            if (_stream.CanSeek)
                _stream.Position = 0;
            var reader = _config.StreamReaderFunc(_stream);
            {
                var row = string.Empty;
                string[] read;
                var firstRow = true;
                Dictionary<int, string> headRows = new Dictionary<int, string>();
                while ((row = reader.ReadLine()) != null)
                {
                    read = Split(row);

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

                        var cell = CustomPropertyHelper.GetEmptyExpandoObject(headRows);
                        for (int i = 0; i <= read.Length - 1; i++)
                            cell[headRows[i]] = read[i];

                        yield return cell;
                        continue;
                    }


                    //body
                    {
                        var cell = CustomPropertyHelper.GetEmptyExpandoObject(read.Length - 1, 0);
                        for (int i = 0; i <= read.Length - 1; i++)
                            cell[ColumnHelper.GetAlphabetColumnName(i)] = read[i];
                        yield return cell;
                    }
                }
            }
        }
        public IEnumerable<T> QueryRange<T>(string sheetName, string startCell, string endCel) where T : class, new()
        {
            return ExcelOpenXmlSheetReader.QueryImplRange<T>(QueryRange(false, sheetName, startCell, endCel), startCell, endCel, this._config);
        }
        public Task<IEnumerable<IDictionary<string, object>>> QueryAsyncRange(bool UseHeaderRow, string sheetName, string startCell, string endCel, CancellationToken cancellationToken = default(CancellationToken))
        {
            return Task.Run(() => QueryRange(UseHeaderRow, sheetName, startCell, endCel), cancellationToken);
        }

        public Task<IEnumerable<T>> QueryAsyncRange<T>(string sheetName, string startCell, string endCel, CancellationToken cancellationToken = default(CancellationToken)) where T : class, new()
        {
            return Task.Run(() => Query<T>(sheetName, startCell), cancellationToken);
        }
        #endregion
    }
}
