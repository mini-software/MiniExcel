using MiniExcelLib.Core;
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
#if NET8_0_OR_GREATER
        var stream = File.Open(templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        await using var disposableStream = stream.ConfigureAwait(false); 
#else
        using var stream = File.Open(templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
#endif
        await SaveAsByTemplateAsync(stream, value, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task SaveAsByTemplateAsync(byte[] templateBytes, object value, CancellationToken cancellationToken = default)
    {
#if NET8_0_OR_GREATER
        var stream = new MemoryStream(templateBytes);
        await using var disposableStream = stream.ConfigureAwait(false); 
#else
        using var stream = new MemoryStream(templateBytes);
#endif
        await SaveAsByTemplateAsync(stream, value, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task SaveAsByTemplateAsync(Stream templateStream, object value, CancellationToken cancellationToken = default)
    {
        if(!templateStream.CanSeek)
            throw new ArgumentException("The template stream must be seekable");
        
        templateStream.Seek(0, SeekOrigin.Begin);
        using var templateReader = await OpenXmlReader.CreateAsync(templateStream, null, cancellationToken: cancellationToken).ConfigureAwait(false);
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

        // Iterate through each entry in the original archive
        foreach (var entry in originalArchive.Entries)
        {
            var entryName = entry.FullName.TrimStart('/');
            if (entryName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) || entryName.Equals("xl/calcChain.xml"))
                continue;
                    
            // Create a new entry in the new archive with the same name
            var newEntry = outputFileArchive.ZipFile.CreateEntry(entry.FullName);

            // Copy the content of the original entry to the new entry
#if NET8_0_OR_GREATER
            var originalEntryStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableEntryStream = originalEntryStream.ConfigureAwait(false);

            var newEntryStream = await newEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableNewEntryStream = newEntryStream.ConfigureAwait(false);
#else
            using var originalEntryStream = entry.Open();
            using var newEntryStream = newEntry.Open();
#endif

            await originalEntryStream.CopyToAsync(newEntryStream
#if NET8_0_OR_GREATER
                , cancellationToken
#endif
            ).ConfigureAwait(false);
        }

        //read sharedString
        var templateSharedStrings = templateReader.SharedStrings;
        templateStream.Position = 0;

        //read all xlsx sheets
        var templateSheets = templateReader.Archive.ZipFile.Entries
            .Where(entry => entry.FullName
                .TrimStart('/')
                .StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase));

        int sheetIdx = 0;
        foreach (var templateSheet in templateSheets)
        {
            // XRowInfos musy be cleared for every sheet or it'll cause duplicates: https://user-images.githubusercontent.com/12729184/115003101-0fcab700-9ed8-11eb-9151-ca4d7b86d59e.png
            _xRowInfos.Clear();
            _xMergeCellInfos.Clear();
            _newXMergeCellInfos.Clear();
            _calcChainCellRefs.Clear();
            
            var templateFullName = templateSheet.FullName;
            var inputValues = _inputValueExtractor.ToValueDictionary(value);
            var outputZipEntry = outputFileArchive.ZipFile.CreateEntry(templateFullName);

#if NET8_0_OR_GREATER
            var outputZipSheetEntryStream = await outputZipEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableSheetEntryStream = outputZipSheetEntryStream.ConfigureAwait(false);
#else
            using var outputZipSheetEntryStream = outputZipEntry.Open();
#endif
            await GenerateSheetByCreateModeAsync(templateSheet, outputZipSheetEntryStream, inputValues, templateSharedStrings, cancellationToken: cancellationToken).ConfigureAwait(false);
            
            //doc.Save(zipStream); //don't do it because: https://user-images.githubusercontent.com/12729184/114361127-61a5d100-9ba8-11eb-9bb9-34f076ee28a2.png
            // disposing writer disposes streams as well. read and parse calc functions before that
            
            sheetIdx++;
            _calcChainContent.Append(CalcChainHelper.GetCalcChainContent(_calcChainCellRefs, sheetIdx));
        }

        // create mode we need to not create first then create here
        var calcChain = outputFileArchive.EntryCollection.FirstOrDefault(e => e.FullName.Contains("xl/calcChain.xml"));
        if (calcChain is not null)
        {
            var calcChainPathName = calcChain.FullName;
            //calcChain.Delete();

            var calcChainEntry = outputFileArchive.ZipFile.CreateEntry(calcChainPathName);
#if NET8_0_OR_GREATER
            var calcChainStream = await calcChainEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableChainEntryStream = calcChainStream.ConfigureAwait(false);
#else
            using var calcChainStream = calcChainEntry.Open();
#endif
            await CalcChainHelper.GenerateCalcChainSheetAsync(calcChainStream, _calcChainContent.ToString(), cancellationToken).ConfigureAwait(false);
        }
        else
        {
            foreach (var entry in originalArchive.Entries)
            {
                if (entry.FullName.Contains("xl/calcChain.xml"))
                {
                    var newEntry = outputFileArchive.ZipFile.CreateEntry(entry.FullName);

                    // Copy the content of the original entry to the new entry
#if NET8_0_OR_GREATER
                    var originalEntryStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
                    await using var disposableEntryStream = originalEntryStream.ConfigureAwait(false);

                    var newEntryStream = await newEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
                    await using var disposableNewEntryStream = newEntryStream.ConfigureAwait(false);
#else
                    using var originalEntryStream = entry.Open();
                    using var newEntryStream = newEntry.Open();
#endif

                    await originalEntryStream.CopyToAsync(newEntryStream
#if NET8_0_OR_GREATER
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
