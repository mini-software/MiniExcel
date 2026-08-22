using MiniExcelLibs.OpenXml;
using System.Globalization;
using System.Text;

namespace MiniExcelLibs.Markdown;

internal static class ExcelMarkdownWriter
{
    private const int DefaultChunkRows = 100;

    internal static async Task WriteAsync(
        Stream xlsxStream,
        Stream markdownStream,
        MarkdownFormat format,
        bool hasHeader,
        string sheetName,
        string sourceName,
        OpenXmlConfiguration configuration,
        CancellationToken cancellationToken)
    {
        if (xlsxStream == null)
            throw new ArgumentNullException(nameof(xlsxStream));
        if (markdownStream == null)
            throw new ArgumentNullException(nameof(markdownStream));
        if (!xlsxStream.CanRead)
            throw new ArgumentException("The XLSX stream must be readable.", nameof(xlsxStream));
        if (!markdownStream.CanWrite)
            throw new ArgumentException("The Markdown stream must be writable.", nameof(markdownStream));
        if (!Enum.IsDefined(typeof(MarkdownFormat), format))
            throw new ArgumentOutOfRangeException(nameof(format));

        cancellationToken.ThrowIfCancellationRequested();
        using var reader = new ExcelOpenXmlSheetReader(xlsxStream, configuration ?? OpenXmlConfiguration.DefaultConfig, leaveOpen: true);
        using var archive = reader._archive;
        var sheets = reader.GetWorkbookRels(reader._archive.entries)
            .Where(sheet => sheetName == null || sheet.Name == sheetName)
            .ToList();

        if (sheetName != null && sheets.Count == 0)
            throw new InvalidOperationException($"Sheet '{sheetName}' does not exist.");

        using var writer = new StreamWriter(markdownStream, new UTF8Encoding(false), 1024, true);
        if (format == MarkdownFormat.LlmFriendly)
        {
            var source = string.IsNullOrEmpty(sourceName) ? "stream" : sourceName;
            await writer.WriteLineAsync($"# {EscapeText(source)}").ConfigureAwait(false);
        }

        for (var sheetIndex = 0; sheetIndex < sheets.Count; sheetIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = sheets[sheetIndex];
            var rows = reader.Query(false, sheet.Name, "A1").GetEnumerator();
            try
            {
                if (!rows.MoveNext())
                {
                    await writer.WriteLineAsync($"## Worksheet: {EscapeText(sheet.Name)}").ConfigureAwait(false);
                    await writer.WriteLineAsync().ConfigureAwait(false);
                    continue;
                }

                var firstRow = rows.Current;
                var columns = firstRow.Keys.ToList();
                var headers = hasHeader
                    ? columns.Select(column => FormatCellValue(firstRow, column)).ToList()
                    : columns.ToList();
                var firstDataRow = hasHeader ? null : firstRow;

                if (format == MarkdownFormat.Simple)
                {
                    if (sheets.Count > 1)
                        await writer.WriteLineAsync($"## Worksheet: {EscapeText(sheet.Name)}").ConfigureAwait(false);
                    await WriteTableStartAsync(writer, headers, columns, false).ConfigureAwait(false);
                    if (firstDataRow != null)
                        await WriteRowAsync(writer, firstDataRow, columns, null).ConfigureAwait(false);

                    while (rows.MoveNext())
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        await WriteRowAsync(writer, rows.Current, columns, null).ConfigureAwait(false);
                    }
                    await writer.WriteLineAsync().ConfigureAwait(false);
                    continue;
                }

                var dataRowNumber = firstDataRow is ExcelOpenXmlSheetReader.ExcelRow firstExcelRow
                    ? firstExcelRow.RowNumber
                    : hasHeader ? 2 : 1;
                var pendingRow = firstDataRow;
                var hasPendingRow = pendingRow != null || rows.MoveNext();
                if (pendingRow == null && hasPendingRow)
                    pendingRow = rows.Current;

                if (!hasPendingRow)
                {
                    var endColumn = columns.Count == 0 ? "A" : columns[columns.Count - 1];
                    await writer.WriteLineAsync($"## Worksheet: {EscapeText(sheet.Name)}").ConfigureAwait(false);
                    await writer.WriteLineAsync($"<!-- miniexcel:chunk range=\"A1:{endColumn}1\" -->").ConfigureAwait(false);
                    await writer.WriteLineAsync().ConfigureAwait(false);
                    await WriteTableStartAsync(writer, headers, columns, true).ConfigureAwait(false);
                    await writer.WriteLineAsync().ConfigureAwait(false);
                    continue;
                }

                while (hasPendingRow)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var chunkStartRow = dataRowNumber;
                    var chunk = new List<MarkdownRow>(DefaultChunkRows);
                    do
                    {
                        var rowNumber = pendingRow is ExcelOpenXmlSheetReader.ExcelRow excelRow
                            ? excelRow.RowNumber
                            : dataRowNumber;
                        chunk.Add(new MarkdownRow(rowNumber, pendingRow));
                        dataRowNumber = rowNumber + 1;
                        pendingRow = null;
                        if (chunk.Count == DefaultChunkRows)
                        {
                            if (rows.MoveNext())
                                pendingRow = rows.Current;
                            break;
                        }
                        if (!rows.MoveNext())
                            break;
                        pendingRow = rows.Current;
                    } while (true);

                    var chunkEndRow = chunk[chunk.Count - 1].RowNumber;
                    var endColumn = columns.Count == 0 ? "A" : columns[columns.Count - 1];
                    await writer.WriteLineAsync($"## Worksheet: {EscapeText(sheet.Name)}").ConfigureAwait(false);
                    await writer.WriteLineAsync($"<!-- miniexcel:chunk range=\"A{chunkStartRow}:{endColumn}{chunkEndRow}\" -->").ConfigureAwait(false);
                    await writer.WriteLineAsync().ConfigureAwait(false);
                    await WriteTableStartAsync(writer, headers, columns, true).ConfigureAwait(false);
                    foreach (var row in chunk)
                        await WriteRowAsync(writer, row.Values, columns, row.RowNumber).ConfigureAwait(false);
                    await writer.WriteLineAsync().ConfigureAwait(false);

                    hasPendingRow = pendingRow != null;
                }
            }
            finally
            {
                rows.Dispose();
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        await writer.FlushAsync().ConfigureAwait(false);
    }

    private static async Task WriteTableStartAsync(StreamWriter writer, IList<string> headers, IList<string> columns, bool includeAddresses)
    {
        var displayHeaders = includeAddresses
            ? columns.Select((column, index) => $"{column}: {headers[index]}").ToList()
            : headers;
        if (includeAddresses)
            displayHeaders.Insert(0, "Row");

        await writer.WriteLineAsync($"| {string.Join(" | ", displayHeaders.Select(EscapeCell))} |").ConfigureAwait(false);
        await writer.WriteLineAsync($"| {string.Join(" | ", displayHeaders.Select(_ => "---"))} |").ConfigureAwait(false);
    }

    private static Task WriteRowAsync(StreamWriter writer, IDictionary<string, object> row, IList<string> columns, int? rowNumber)
    {
        var values = columns.Select(column => EscapeCell(FormatCellValue(row, column))).ToList();
        if (rowNumber.HasValue)
            values.Insert(0, rowNumber.Value.ToString(CultureInfo.InvariantCulture));
        return writer.WriteLineAsync($"| {string.Join(" | ", values)} |");
    }

    private static string FormatCellValue(IDictionary<string, object> row, string column)
    {
        var value = FormatValue(row[column]);
        if (row is ExcelOpenXmlSheetReader.ExcelRow excelRow && excelRow.Formulas.TryGetValue(column, out var formula))
            return string.IsNullOrEmpty(value) ? $"={formula}" : $"{value} (formula: ={formula})";
        return value;
    }

    private static string FormatValue(object value)
    {
        if (value == null)
            return string.Empty;
        if (value is IFormattable formattable)
            return formattable.ToString(null, CultureInfo.InvariantCulture);
        return value.ToString();
    }

    private static string EscapeText(string value) => EscapeCell(value);

    private static string EscapeCell(string value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;

        var escaped = value.Replace("\\", "\\\\");
        foreach (var character in new[] { "`", "*", "_", "{", "}", "[", "]", "(", ")", "#", "+", "-", "!", "|", "~" })
            escaped = escaped.Replace(character, "\\" + character);

        escaped = escaped
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\r\n", "<br>")
            .Replace("\r", "<br>")
            .Replace("\n", "<br>");

        return escaped;
    }

    private sealed class MarkdownRow
    {
        internal MarkdownRow(int rowNumber, IDictionary<string, object> values)
        {
            RowNumber = rowNumber;
            Values = values;
        }

        internal int RowNumber { get; }
        internal IDictionary<string, object> Values { get; }
    }
}