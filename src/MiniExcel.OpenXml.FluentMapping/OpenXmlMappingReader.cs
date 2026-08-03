using MiniExcelLib.OpenXml.Reader;
using MiniExcelLib.OpenXml.Styles;

namespace MiniExcelLib.OpenXml.FluentMapping;

internal partial class OpenXmlMappingReader(OpenXmlZip archive, IMiniExcelConfiguration? configuration) : OpenXmlReader(archive, configuration)
{
    [CreateSyncVersion]
    internal new static async Task<OpenXmlMappingReader> CreateAsync(Stream stream, IMiniExcelConfiguration? configuration, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        OpenXmlZip? archive = null;
        OpenXmlMappingReader? reader = null;

        try
        {
            ThrowHelper.ThrowIfInvalidOpenXml(stream);
    
            archive = await OpenXmlZip.CreateAsync(stream, leaveOpen: leaveOpen, cancellationToken: cancellationToken).ConfigureAwait(false);
            reader = new OpenXmlMappingReader(archive, configuration);
            await reader.SetSharedStringsAsync(cancellationToken).ConfigureAwait(false);

            var result = reader;
            reader = null;
            archive = null;
            stream = null!;
            
            return result;
        }
        finally
        {
#if SYNC_ONLY
            reader?.Dispose();
#else
            if (reader?.DisposeAsync() is { } disposeTask)
                await disposeTask.ConfigureAwait(false);
#endif

            if (archive is not null)
                await archive.DisposeAsync().ConfigureAwait(false);
            
            if (!leaveOpen && (Stream?)stream is not null)
                await stream.DisposeAsync().ConfigureAwait(false);
        }
    }
    
    /// <summary>
    /// Direct mapped query that bypasses dictionary creation for better performance
    /// </summary>
    [CreateSyncVersion]
    internal async IAsyncEnumerable<MappedRow> QueryMappedAsync(string? sheetName, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        
        const bool withoutCr = false;
        var sheetEntry = GetSheetEntry(sheetName);

        MergeCells? mergeCells = null;
        if (_config.FillMergedCells)
        {
            var mergeCellsResult = await TryGetMergeCellsAsync(sheetEntry, cancellationToken).ConfigureAwait(false); 
            if (mergeCellsResult.Success)
                mergeCells = mergeCellsResult.MergeCells;
        }
        
        // Direct XML reading without dictionary creation
        var xmlSettings = XmlReaderHelper.GetXmlReaderSettings();

        var sheetStream = await sheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableSheetStream = sheetStream.ConfigureAwait(false);

        using var reader = XmlReader.Create(sheetStream, xmlSettings);
        
        if (!reader.IsStartElement("worksheet", Ns))
            yield break;

        if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
            yield break;

        while (!reader.EOF)
        {
            if (reader.IsStartElement("sheetData", Ns))
            {
                if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                    continue;

                int rowIndex = -1;
                while (!reader.EOF)
                {
                    if (reader.IsStartElement("row", Ns))
                    {
                        if (int.TryParse(reader.GetAttribute("r"), out int arValue))
                            rowIndex = arValue - 1; // The row attribute is 1-based
                        else
                            rowIndex++;

                        // Read row directly into mapped structure
                        await foreach (var mappedRow in ReadMappedRowAsync(reader, rowIndex, withoutCr, mergeCells, cancellationToken).ConfigureAwait(false))
                        {
                            yield return mappedRow;
                        }
                    }
                    else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                    {
                        break;
                    }
                }
            }
            else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
            {
                break;
            }
        }
    }
    
    [CreateSyncVersion]
    private async IAsyncEnumerable<MappedRow> ReadMappedRowAsync(
        XmlReader reader,
        int rowIndex,
        bool withoutCr,
        MergeCells? mergeCells,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
        {
            // Empty row
            yield return new MappedRow(rowIndex);
            yield break;
        }

        var row = new MappedRow(rowIndex);
        var columnIndex = withoutCr ? -1 : 0;
        
        while (!reader.EOF)
        {
            if (reader.IsStartElement("c", Ns))
            {
                var aS = reader.GetAttribute("s");
                var aR = reader.GetAttribute("r");
                var aT = reader.GetAttribute("t");
                
                var cellAndColumn = await ReadCellAndSetColumnIndexAsync(reader, columnIndex, withoutCr, 0, aR, aT, cancellationToken).ConfigureAwait(false);
                var cellValue = cellAndColumn.CellValue;
                columnIndex = cellAndColumn.ColumnIndex;

                if (_config.FillMergedCells && mergeCells is not null)
                {
                    if (mergeCells.MergesValues.ContainsKey(aR))
                    {
                        mergeCells.MergesValues[aR] = cellValue;
                    }
                    else if (mergeCells.MergesMap.TryGetValue(aR, out var mergeKey))
                    {
                        mergeCells.MergesValues.TryGetValue(mergeKey, out cellValue);
                    }
                }

                if (!string.IsNullOrEmpty(aS)) // Custom style
                {
                    if (int.TryParse(aS, NumberStyles.Any, CultureInfo.InvariantCulture, out var styleIndex))
                    {
                        _style ??= await OpenXmlStyles.CreateAsync(Archive, cancellationToken).ConfigureAwait(false);
                        cellValue = _style.ConvertValueByStyleFormat(styleIndex, cellValue);
                    }
                }

                row.SetCell(columnIndex, cellValue);
            }
            else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
            {
                break;
            }
        }
        
        yield return row;
    }
}
