using MiniExcelLib.Exceptions;
using MiniExcelLib.Helpers;
using MiniExcelLib.Reflection;
using IMiniExcelReader = MiniExcelLib.Abstractions.IMiniExcelReader;
using MiniExcelMapper = MiniExcelLib.Reflection.MiniExcelMapper;

namespace MiniExcelLib.Csv;

internal partial class CsvReader : IMiniExcelReader
{
    private readonly Stream _stream;
    private readonly CsvConfiguration _config;

    internal CsvReader(Stream stream, IMiniExcelConfiguration? configuration)
    {
        _stream = stream;
        _config = configuration as CsvConfiguration ?? CsvConfiguration.DefaultConfiguration;
    }

    [CreateSyncVersion]
    public async IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(bool useHeaderRow, string? sheetName, string startCell, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (startCell != "A1")
            throw new NotImplementedException("CSV does not implement parameter startCell");

        if (_stream.CanSeek)
            _stream.Position = 0;

        var reader = _config.StreamReaderFunc(_stream);
        var firstRow = true;
        var headRows = new Dictionary<int, string>();

        var rowIndex = 0;
        while (await reader.ReadLineAsync(
#if NET7_0_OR_GREATER
            cancellationToken
#endif
            ).ConfigureAwait(false) is { } row)
        {
            rowIndex++;
            string finalRow = row;
            
            if (_config.ReadLineBreaksWithinQuotes)
            {
                while (finalRow.Count(c => c == '"') % 2 != 0)
                {
                    var nextPart = await reader.ReadLineAsync(
#if NET7_0_OR_GREATER
                        cancellationToken
#endif
                    ).ConfigureAwait(false);
                    
                    if (nextPart is null)
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

                throw new MiniExcelColumnNotFoundException(columnIndex: null, headRows[colIndex], [], rowIndex, headers, rowValues, $"Csv read error: Column {colIndex} not found in Row {rowIndex}");
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

            // todo: can we find a way to remove the redundant cell conversions for CSV?
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

    [CreateSyncVersion]
    public IAsyncEnumerable<T> QueryAsync<T>(string? sheetName, string startCell, bool mapHeaderAsData, CancellationToken cancellationToken = default) where T : class, new()
    {
        var records = QueryAsync(false, sheetName, startCell, cancellationToken);
        return MiniExcelMapper.MapQueryAsync<T>(records, startCell, mapHeaderAsData, false, _config, cancellationToken);
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(bool useHeaderRow, string? sheetName, string startCell, string endCell, CancellationToken cancellationToken = default)
    {
        throw new NotImplementedException("CSV does not implement QueryRange");
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<T> QueryRangeAsync<T>(string? sheetName, string startCell, string endCell, bool treatHeaderAsData, CancellationToken cancellationToken = default) where T : class, new()
    {
        var dynamicRecords = QueryRangeAsync(false, sheetName, startCell, endCell, cancellationToken);
        return MiniExcelMapper.MapQueryAsync<T>(dynamicRecords, startCell, treatHeaderAsData, false, _config, cancellationToken);
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(bool useHeaderRow, string? sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, CancellationToken cancellationToken = default)
    {
        throw new NotImplementedException("CSV does not implement QueryRange");
    }

    [CreateSyncVersion]
    public IAsyncEnumerable<T> QueryRangeAsync<T>(string? sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, bool treatHeaderAsData, CancellationToken cancellationToken = default) where T : class, new()
    {
        var dynamicRecords = QueryRangeAsync(false, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken);
        return MiniExcelMapper.MapQueryAsync<T>(dynamicRecords, ConvertXyToCell(startRowIndex, startColumnIndex), treatHeaderAsData, false, _config, cancellationToken);
    }

    private static string ConvertXyToCell(int x, int y)
    {
        int dividend = x;
        string columnName = string.Empty;

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return $"{columnName}{y}";
    }
    
    private string[] Split(string row)
    {
        if (_config.SplitFn is not null)
            return _config.SplitFn(row);
        
        //this code from S.O : https://stackoverflow.com/a/11365961/9131476
        return Regex.Split(row, $"[\t{_config.Seperator}](?=(?:[^\"]|\"[^\"]*\")*$)")
            .Select(s => Regex.Replace(s.Replace("\"\"", "\""), "^\"|\"$", ""))
            .ToArray();
    }

    public void Dispose()
    {
        _stream?.Dispose();
    }
}