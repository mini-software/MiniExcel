namespace MiniExcelLib.OpenXml.Reader;

internal partial class OpenXmlReader
{
    [CreateSyncVersion]
    internal IAsyncEnumerable<T> QueryTableAsync<T>(string sheetName, string tableName, CancellationToken cancellationToken = default)
        where T : class, new()
    {
        var query = QueryTableAsync(sheetName, tableName, true, cancellationToken);
        return MiniExcelMapper.MapQueryAsync<T>(query, 0, false, _config.TrimColumnNames, _config, XmlHelper.DecodeString, cancellationToken);    
    }

    [CreateSyncVersion]
    internal async IAsyncEnumerable<IDictionary<string, object?>> QueryTableAsync(string sheetName, string tableName, bool prependHeaders, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        TableInfo? table = null;
        await foreach (var item in GetTableInfosAsync(sheetName, cancellationToken).ConfigureAwait(false))
        {
            if (item.Name.Equals(tableName, StringComparison.OrdinalIgnoreCase))
            {
                table = item;
                break;
            }
        }

        if (table is null)
            throw new InvalidDataException($"The table {tableName} was not found.");

        if (table.ReferenceCells?.Split(':') is not [var start, var end] ||
            !CellReferenceConverter.TryParseCellReference(start, out var startCol, out var startRow) ||
            !CellReferenceConverter.TryParseCellReference(end, out var endCol, out var endRow))
        {
            throw new InvalidDataException("A valid cell range could not be extracted from the table metadata.");
        }

        if (!table.HiddenHeader)
            startRow++;

        if (prependHeaders)
        {
            var headers = ExpandoHelper.CreateEmptyByIndices(endCol - 1, startCol - 1);
            var columnCount = Math.Min(headers.Count, table.Columns.Length);

            for (int i = 0; i < columnCount; i++)
            {
                var index = CellReferenceConverter.GetAlphabeticalIndex(startCol + i - 1);
                headers[index] = table.Columns[i];
            }

            yield return headers;
        }
        
        await foreach (var row in QueryRangeAsync(false, sheetName, startRow, startCol, endRow, endCol, cancellationToken).ConfigureAwait(false))
        {
            if (!prependHeaders)
            {
                for (var i = 0; i < table.Columns.Length; i++)
                {
                    var oldHeader = CellReferenceConverter.GetAlphabeticalIndex(i + startCol - 1);
                    if (row.TryGetValue(oldHeader, out var cellValue))
                    {
                        var newHeader = table.Columns[i];
                        row[newHeader] = cellValue;
                        if (newHeader != oldHeader)
                        {
                            row.Remove(oldHeader);
                        }
                    }
                }
            }

            yield return row;
        }
    }

    [CreateSyncVersion]
    private async IAsyncEnumerable<TableInfo> GetTableInfosAsync(string sheetName, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var rels = await GetWorkbookRelsAsync(Archive.EntryCollection, cancellationToken).ConfigureAwait(false);
        if (rels?.Find(x => x.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)) is not { Path: { } path })
            throw new InvalidDataException($"Worksheet {sheetName} was not found.");
        
        List<string> tables = [];
        var sheetFilename = path.Split('/')[^1];

        if (Archive.GetEntry($"xl/worksheets/_rels/{sheetFilename}.rels") is { } entry)
        {
            var entryStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableEntryStream = entryStream.ConfigureAwait(false); 

            var readerSettings = XmlReaderHelper.GetXmlReaderSettings();
            using var reader = XmlReader.Create(entryStream, readerSettings);

            if (!reader.ReadToFollowing("Relationship"))
                yield break;

            do
            {
                if (reader.GetAttribute("Type") == Schemas.SpreadsheetmlXmlTableRelationship)
                {
                    if (reader.GetAttribute("Target") is { } target &&
                        target.Split('/').LastOrDefault() is { } table)
                    {
                        tables.Add(table);
                    }
                }
            }
            while(reader.ReadToNextSibling("Relationship"));
        }

        foreach (var table in tables)
        {
            if (Archive.GetEntry($"xl/tables/{table}") is not { } tableEntry)
                continue;

            var entryStream = await tableEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableEntryStream = entryStream.ConfigureAwait(false);
            using var reader = XmlReader.Create(entryStream, XmlReaderHelper.GetXmlReaderSettings());

            if (!reader.ReadToFollowing("table"))
                continue;

            if (reader.GetAttribute("name") is not { } tableName || 
                reader.GetAttribute("ref") is not {  } @ref)
            {
                continue;
            }

            var headerIsHidden = reader.GetAttribute("headerRowCount") == "0";
            if (!reader.ReadToDescendant("tableColumn"))
                continue;

            List<string> columns = [];
            var colCount = 0;

            do
            {
                var colName = reader.GetAttribute("name") ?? $"Column{colCount}";
                columns.Add(colName);
                colCount++;
            }
            while (reader.ReadToNextSibling("tableColumn"));
                
            yield return new TableInfo(tableName, [..columns], @ref, headerIsHidden);
        }
    }
}
