using MiniExcelLib.OpenXml.Reader;
using CalcChainHelper = MiniExcelLib.OpenXml.Utils.CalcChainHelper;

namespace MiniExcelLib.OpenXml.Templates;

internal partial class OpenXmlTemplate : IMiniExcelTemplate
{
    private readonly Stream _outputFileStream;
    private readonly OpenXmlConfiguration _configuration;
    private readonly OpenXmlValueExtractor _inputValueExtractor;
    private readonly StringBuilder _calcChainContent = new();

    internal OpenXmlTemplate(Stream stream, IMiniExcelConfiguration? configuration, OpenXmlValueExtractor inputValueExtractor)
    {
        _outputFileStream = stream;
        _configuration = (OpenXmlConfiguration?)configuration ?? OpenXmlConfiguration.Default;
        _inputValueExtractor = inputValueExtractor;
    }

    [CreateSyncVersion]
    public async Task SaveAsByTemplateAsync(string templatePath, object value, CancellationToken cancellationToken = default)
    {
        var stream = File.Open(templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        await using var disposableStream = stream.ConfigureAwait(false); 

        await SaveAsByTemplateAsync(stream, value, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task SaveAsByTemplateAsync(byte[] templateBytes, object value, CancellationToken cancellationToken = default)
    {
        var stream = new MemoryStream(templateBytes);
        await using var disposableStream = stream.ConfigureAwait(false); 

        await SaveAsByTemplateAsync(stream, value, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task SaveAsByTemplateAsync(Stream templateStream, object value, CancellationToken cancellationToken = default)
    {
        if (!templateStream.CanSeek)
            throw new ArgumentException("The template stream must be seekable");

        templateStream.Seek(0, SeekOrigin.Begin);
        var templateReader = await OpenXmlReader.CreateAsync(templateStream, null, cancellationToken: cancellationToken).ConfigureAwait(false);
        await using var disposableTemplateReader = templateReader.ConfigureAwait(false);

        var outputFileArchive = await OpenXmlZip.CreateAsync(_outputFileStream, mode: ZipArchiveMode.Create, true, Encoding.UTF8, isUpdateMode: false, cancellationToken: cancellationToken).ConfigureAwait(false);
        await using var disposableOutputFileArchive = outputFileArchive.ConfigureAwait(false);

        try
        {
            outputFileArchive.EntryCollection = templateReader.Archive.ZipFile.Entries; //TODO:need to remove
        }
        catch (InvalidDataException e)
        {
            throw new InvalidDataException($"An invalid OpenXml zip archive was detected, please check or open an issue for this error: {e.Message}");
        }

        // foreach all templateStream and create file for _outputFileStream and not create sheet file
        foreach (var entry in templateReader.Archive.ZipFile.Entries)
        {
            outputFileArchive.Entries.Add(entry.FullName.Replace('\\', '/'), entry);
        }

        // Create a new zip file for writing
        templateStream.Position = 0;
#if NET10_0_OR_GREATER
        var originalArchive = await ZipArchive.CreateAsync(templateStream, ZipArchiveMode.Read, false, null, cancellationToken).ConfigureAwait(false);
        await using var disposableArchive = originalArchive.ConfigureAwait(false);
#else
        using var originalArchive = new ZipArchive(templateStream, ZipArchiveMode.Read);
#endif
        // sheet name map
        var sheetPathRealNameMap = await GetSheetNameMapAsync(originalArchive, cancellationToken).ConfigureAwait(false);

        // Iterate through each entry in the original archive
        foreach (var entry in originalArchive.Entries)
        {
            var entryName = entry.FullName.TrimStart('/');
            if (entryName.StartsWith(ExcelFileNames.WorksheetBase, StringComparison.OrdinalIgnoreCase) || 
                entryName.Equals(ExcelFileNames.CalcChain, StringComparison.OrdinalIgnoreCase) ||
                entryName.Equals(ExcelFileNames.Workbook, StringComparison.OrdinalIgnoreCase) ||
                entryName.Equals(ExcelFileNames.WorkbookRels, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            // Create a new entry in the new archive with the same name
            var newEntry = outputFileArchive.ZipFile.CreateEntry(entry.FullName);

            // Copy the content of the original entry to the new entry
            var originalEntryStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableEntryStream = originalEntryStream.ConfigureAwait(false);

            var newEntryStream = await newEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableNewEntryStream = newEntryStream.ConfigureAwait(false);

            await originalEntryStream.CopyToAsync(newEntryStream
#if NET
                , cancellationToken
#endif
            ).ConfigureAwait(false);
        }

        //read sharedString
        var templateSharedStrings = templateReader.SharedStrings;
        templateStream.Position = 0;

        //read all xlsx sheets
        var templateSheets = templateReader.Archive.ZipFile.Entries
            .Where(entry => entry.FullName.TrimStart('/').StartsWith(ExcelFileNames.WorksheetBase, StringComparison.OrdinalIgnoreCase));

        int sheetIdx = 0;
        // collect all sheet info for batch add to config, avoid duplicated and missing sheet name when create mode
        List<(int Index, string Name)> allSheetInfos = [];

        foreach (var templateSheet in templateSheets)
        {
            // XRowInfos musy be cleared for every sheet or it'll cause duplicates: https://user-images.githubusercontent.com/12729184/115003101-0fcab700-9ed8-11eb-9151-ca4d7b86d59e.png
            _xRowInfos.Clear();
            _xMergeCellInfos.Clear();
            _newXMergeCellInfos.Clear();
            _calcChainCellRefs.Clear();

            var templateFullName = templateSheet.FullName;
            var inputValues = _inputValueExtractor.ToValueDictionary(value);
            sheetPathRealNameMap.TryGetValue(templateFullName, out var sheetName);

            if (await TryExpandParametrizedSheetAsync(outputFileArchive, sheetName, templateSharedStrings, sheetIdx, allSheetInfos, templateSheet, inputValues, cancellationToken).ConfigureAwait(false))
                break;

            var outputZipEntry = outputFileArchive.ZipFile.CreateEntry(templateFullName);
            var outputZipSheetEntryStream = await outputZipEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableSheetEntryStream = outputZipSheetEntryStream.ConfigureAwait(false);

            await GenerateSheetByCreateModeAsync(templateSheet, outputZipSheetEntryStream, inputValues, templateSharedStrings, cancellationToken: cancellationToken).ConfigureAwait(false);
            // disposing writer disposes streams as well, read and parse calc functions before that

            sheetIdx++;
            allSheetInfos.Add((sheetIdx, sheetName));
            _calcChainContent.Append(CalcChainHelper.GetCalcChainContent(_calcChainCellRefs, sheetIdx));
        }

        // batch add sheet
        await BatchAddSheetsToWorkbookAsync(outputFileArchive.ZipFile, originalArchive, allSheetInfos, cancellationToken).ConfigureAwait(false);

        // create mode we need to not create first then create here
        var calcChain = outputFileArchive.EntryCollection.FirstOrDefault(e 
            => e.FullName.TrimStart('/').Equals(ExcelFileNames.CalcChain, StringComparison.OrdinalIgnoreCase));

        if (calcChain is not null)
        {
            var calcChainEntry = outputFileArchive.ZipFile.CreateEntry(calcChain.FullName);
            var calcChainStream = await calcChainEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableChainEntryStream = calcChainStream.ConfigureAwait(false);

            await CalcChainHelper.GenerateCalcChainSheetAsync(calcChainStream, _calcChainContent.ToString(), cancellationToken).ConfigureAwait(false);
        }
        else
        {
            foreach (var entry in originalArchive.Entries)
            {
                if (entry.FullName.TrimStart('/').Equals(ExcelFileNames.CalcChain, StringComparison.OrdinalIgnoreCase))
                {
                    var newEntry = outputFileArchive.ZipFile.CreateEntry(entry.FullName);

                    // Copy the content of the original entry to the new entry
                    var originalEntryStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
                    await using var disposableEntryStream = originalEntryStream.ConfigureAwait(false);

                    var newEntryStream = await newEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
                    await using var disposableNewEntryStream = newEntryStream.ConfigureAwait(false);

                    await originalEntryStream.CopyToAsync(newEntryStream
#if NET
                        , cancellationToken
#endif
                    ).ConfigureAwait(false);
                }
            }
        }

#if NET10_0_OR_GREATER
        await outputFileArchive.ZipFile.DisposeAsync().ConfigureAwait(false);
#else
        outputFileArchive.ZipFile.Dispose();
#endif
    }
}
