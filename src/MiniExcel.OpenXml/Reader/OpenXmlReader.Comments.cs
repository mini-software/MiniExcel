namespace MiniExcelLib.OpenXml.Reader;

internal partial class OpenXmlReader
{
    private static readonly XNamespace NsRel = Schemas.OpenXmlPackageRelationships;
    private static readonly XNamespace Ns18Tc = Schemas.SpreadsheetmlXmlX18Tc;
    private static readonly XNamespace NsMain = Schemas.SpreadsheetmlXmlMain;
    private static readonly XNamespace Ns14R = Schemas.SpreadsheetmlXmlX14R;
        
    [CreateSyncVersion]
    internal async Task<CommentResultSet> ReadCommentsAsync(string? sheetName, CancellationToken cancellationToken = default)
    {
        SetWorkbookRels(Archive.EntryCollection);

        var sheetRecord = string.IsNullOrEmpty(sheetName) 
            ? _sheetRecords?.FirstOrDefault() 
            : _sheetRecords?.SingleOrDefault(s => s.Name.Equals(sheetName, StringComparison.CurrentCultureIgnoreCase));

        if (sheetRecord?.Path?.Split('/')[^1] is not { } sheetFile)
            throw new InvalidDataException("A valid worksheet could not be found.");

        if (string.IsNullOrEmpty(sheetName))
            sheetName = sheetRecord.Name;

        if (Archive.GetEntry($"xl/worksheets/_rels/{sheetFile}.rels") is not { } rel)
            return new CommentResultSet(sheetName, [], []);

        var stream = await rel.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableStream = stream.ConfigureAwait(false);  

        var relDoc = await XDocument.LoadAsync(stream, LoadOptions.None, cancellationToken).ConfigureAwait(false);
        HashSet<string?> refCells = [];

        var people = await GetAuthorsAsync(cancellationToken).ConfigureAwait(false);
        var commentThreads = await GetThreadedCommentsAsync(relDoc, refCells, people, cancellationToken).ConfigureAwait(false);
        var notes = await GetNotesAsync(relDoc, refCells, cancellationToken).ConfigureAwait(false);

        return new CommentResultSet(sheetName, commentThreads, notes);
    }

    [CreateSyncVersion]
    private async Task<ICollection<Author>> GetAuthorsAsync(CancellationToken cancellationToken)
    {
        if (Archive.GetEntry(ExcelFileNames.Person) is not { } persons)
            return [];

        var personStream = await persons.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposablePersonStream = personStream.ConfigureAwait(false);  
            
        var personDoc = await XDocument.LoadAsync(personStream, LoadOptions.None, cancellationToken).ConfigureAwait(false);
        var personElements = personDoc.Root?.Elements(Ns18Tc + "person");

        return personElements
            ?.Select(p => new Author
            {
                Id = Guid.Parse(p.Attribute("id")!.Value),
                DisplayName = p.Attribute("displayName")?.Value is { } name and not "" ? name : "???",
                ProviderId = p.Attribute("providerId")?.Value,
            })
            .ToList() ?? [];
    }

    [CreateSyncVersion]
    private async Task<List<NoteComment>> GetNotesAsync(XDocument relDoc, HashSet<string?> refCells, CancellationToken cancellationToken)
    {
        var noteRels = relDoc.Root?.Elements(NsRel + "Relationship");
        var notesElement = noteRels?.FirstOrDefault(x => x.Attribute("Type")?.Value == Schemas.SpreadsheetmlXmlCommentsRelationship);
        var notesTarget = notesElement?.Attribute("Target");
        var notesPath = notesTarget?.Value.TrimStart('.', '/');

        if (Archive.GetEntry($"xl/{notesPath}") is not { } noteEntry)
            return [];

        var noteEntryStream = await noteEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableNoteEntryStream = noteEntryStream.ConfigureAwait(false);

        var doc = await XDocument.LoadAsync(noteEntryStream, LoadOptions.None, cancellationToken).ConfigureAwait(false);

        var authorElements = doc.Root?.Element(NsMain + "authors")?.Elements(NsMain + "author");
        var authors = authorElements?.Select(a => a.Value).ToArray();

        var commentElements = doc.Root
            ?.Element(NsMain + "commentList")
            ?.Elements(NsMain + "comment");

        return commentElements
            ?.Where(c => !refCells.Contains(c.Attribute("ref")?.Value))
            .Select(c => new NoteComment
            {
                Id = Guid.TryParse(c.Attribute(Ns14R + "uid")?.Value.Trim('{', '}'), out var noteId) ? noteId : Guid.Empty,
                Author = int.TryParse(c.Attribute("authorId")?.Value, out var authorId) ? authors?.ElementAtOrDefault(authorId) : "",
                ReferenceCell =  c.Attribute("ref")?.Value,
                Text = string.Join("", GetTextFromComment(c))
            })
            .ToList() ?? [];
    }

    [CreateSyncVersion]
    private async Task<List<ThreadedComment>> GetThreadedCommentsAsync(XDocument relDoc, HashSet<string?> refCells, ICollection<Author> people, CancellationToken cancellationToken)
    {
        var threadedCommentRels = relDoc.Root?.Elements(NsRel + "Relationship");
        var threadedCommentsElement = threadedCommentRels?.FirstOrDefault(x => x.Attribute("Type")?.Value == Schemas.SpreadsheetmlXmlThreadedCommentRelationship);
        var threadedCommentsTarget = threadedCommentsElement?.Attribute("Target");
        var threadedCommentsPath = threadedCommentsTarget?.Value.TrimStart('.', '/');

        if (Archive.GetEntry($"xl/{threadedCommentsPath}") is not { } threadEntry)
            return [];

        var threadEntryStream = await threadEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableThreadEntryStream = threadEntryStream.ConfigureAwait(false);

        var doc = await XDocument.LoadAsync(threadEntryStream, LoadOptions.None, cancellationToken).ConfigureAwait(false);

        var commentThreadElements = doc.Root?.Elements(Ns18Tc + "threadedComment");
        var commentThreads = commentThreadElements
            ?.Where(tc => tc.Attribute("parentId") is null)
            .Select(tc => new ThreadedComment
            {
                Id = Guid.Parse(tc.Attribute("id")!.Value.Trim('{', '}')),
                Author = people.FirstOrDefault(p => p.Id == (Guid.TryParse(tc.Attribute("personId")?.Value, out var person) ? person : Guid.Empty)),
                CreatedAt = DateTime.Parse(tc.Attribute("dT")!.Value, CultureInfo.InvariantCulture),
                ReferenceCell = tc.Attribute("ref")?.Value!,
                Text = tc.Value,
                Resolved = tc.Attribute("done")?.Value is not (null or "0")
            })
            .ToList() ?? [];

        var replyElements = doc.Root?.Elements(Ns18Tc + "threadedComment");
        var replies = replyElements
            ?.Where(tc => tc.Attribute("parentId") is not null)
            .Select(tc => new ThreadedCommentReply
            {
                Id = Guid.Parse(tc.Attribute("id")!.Value.Trim('{', '}')),
                ParentId = Guid.Parse(tc.Attribute("parentId")!.Value),
                Author = people.FirstOrDefault(p => p.Id == Guid.Parse(tc.Attribute("personId")!.Value)),
                CreatedAt = DateTime.Parse(tc.Attribute("dT")!.Value, CultureInfo.InvariantCulture),
                Text = tc.Value
            })
            .ToLookup(x => x.ParentId);

        foreach (var thread in commentThreads)
        {
            refCells.Add(thread.ReferenceCell);

            if (replies is not null)
                thread.ThreadedComments = replies[thread.Id].ToList();
        }

        return commentThreads;
    }
    
    private static IEnumerable<string?> GetTextFromComment(XElement? comment)
    {
        return comment?.Element(NsMain + "text") is { } textElement
            ? textElement.Descendants(NsMain + "t").Select(t => t.Value)
            : [];
    } 
}
