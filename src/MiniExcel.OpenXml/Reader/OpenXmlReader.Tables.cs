namespace MiniExcelLib.OpenXml.Reader;

internal partial class OpenXmlReader
{
    [CreateSyncVersion]
    internal IAsyncEnumerable<T> QueryTableAsync<T>(string sheetName, string tableName, CancellationToken cancellationToken = default)
        where T : class, new()
    {
        var query = QueryTableAsync(sheetName, tableName, cancellationToken);
        return MiniExcelMapper.MapQueryAsync<T>(query, 0, false, _config.TrimColumnNames, _config, XmlHelper.DecodeString, cancellationToken);    
    }

    [CreateSyncVersion]
    internal async IAsyncEnumerable<IDictionary<string, object?>> QueryTableAsync(string sheetName, string tableName, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        string? refCells = null;
        await foreach (var item in GetTableInfosAsync(sheetName, cancellationToken).ConfigureAwait(false))
        {
            if (item.Name.Equals(tableName, StringComparison.OrdinalIgnoreCase))
            {
                refCells = item.Ref;
                break;
            }
        }

        if (refCells is null)
            throw new InvalidDataException($"The table {tableName} was not found.");

        if (refCells.Split(':') is not [var start, var end] ||
            !CellReferenceConverter.TryParseCellReference(start, out _, out _) ||
            !CellReferenceConverter.TryParseCellReference(end, out _, out _))
        {
            throw new InvalidDataException("A valid cell range could not be extracted from the table metadata.");
        }

        await foreach (var row in QueryRangeAsync(false, sheetName, start, end, cancellationToken).ConfigureAwait(false))
            yield return row;
    }
    
    [CreateSyncVersion]
    private async IAsyncEnumerable<(string Name, string Ref)> GetTableInfosAsync(string sheetName, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        var rels = await GetWorkbookRelsAsync(Archive.EntryCollection, cancellationToken).ConfigureAwait(false);
        if (rels?.Find(x => x.Name == sheetName) is not { Path: { } path })
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
            if (Archive.GetEntry($"xl/tables/{table}") is { } tableEntry)
            {
                var entryStream = await tableEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
                await using var disposableEntryStream = entryStream.ConfigureAwait(false);
                using var reader = XmlReader.Create(entryStream, XmlReaderHelper.GetXmlReaderSettings());

                reader.ReadToFollowing("table");
                var name = reader.GetAttribute("name")!;
                var @ref = reader.GetAttribute("ref")!;
 
                yield return (name, @ref);
            }
        }
    }
}
