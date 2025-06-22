﻿using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.OpenXml.Constants;
using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using Zomp.SyncMethodGenerator;

namespace MiniExcelLibs.SaveByTemplate;

internal partial class ExcelOpenXmlTemplate : IExcelTemplate
{
#if NET7_0_OR_GREATER
    [GeneratedRegex("(?<={{).*?(?=}})")] private static partial Regex ExpressionRegex();
    private static readonly Regex IsExpressionRegex = ExpressionRegex();
#else
    private static readonly Regex IsExpressionRegex = new("(?<={{).*?(?=}})");
#endif 
    private static readonly XmlNamespaceManager Ns;

    private readonly Stream _outputFileStream;
    private readonly OpenXmlConfiguration _configuration;
    private readonly IInputValueExtractor _inputValueExtractor;
    private readonly StringBuilder _calcChainContent = new();

    static ExcelOpenXmlTemplate()
    {
        Ns = new XmlNamespaceManager(new NameTable());
        Ns.AddNamespace("x", Schemas.SpreadsheetmlXmlns);
        Ns.AddNamespace("x14ac", Schemas.SpreadsheetmlXmlX14Ac);
    }

    public ExcelOpenXmlTemplate(Stream stream, IMiniExcelConfiguration? configuration, InputValueExtractor inputValueExtractor)
    {
        _outputFileStream = stream;
        _configuration = (OpenXmlConfiguration?)configuration ?? OpenXmlConfiguration.DefaultConfig;
        _inputValueExtractor = inputValueExtractor;
    }

    [CreateSyncVersion]
    public async Task SaveAsByTemplateAsync(string templatePath, object value, CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(templatePath);
        await SaveAsByTemplateImplAsync(stream, value, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task SaveAsByTemplateAsync(byte[] templateBytes, object value, CancellationToken cancellationToken = default)
    {
        using Stream stream = new MemoryStream(templateBytes);
        await SaveAsByTemplateImplAsync(stream, value, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    internal async Task SaveAsByTemplateImplAsync(Stream templateStream, object value, CancellationToken cancellationToken = default)
    {
        // foreach all templateStream and create file for _outputFileStream and not create sheet file
        templateStream.Position = 0;
        using var templateReader = await ExcelOpenXmlSheetReader.CreateAsync(templateStream, null, cancellationToken: cancellationToken).ConfigureAwait(false);
        using var outputFileArchive = new ExcelOpenXmlZip(_outputFileStream, mode: ZipArchiveMode.Create, true, Encoding.UTF8, isUpdateMode: false);
        try
        {
            outputFileArchive.EntryCollection = templateReader.Archive.ZipFile.Entries; //TODO:need to remove
        }
        catch (InvalidDataException e)
        {
            throw new InvalidDataException($"An invalid valid OpenXml zip archive was detected, please check or open an issue for this error: {e.Message}");
        }

        foreach (var entry in templateReader.Archive.ZipFile.Entries)
        {
            outputFileArchive.Entries.Add(entry.FullName.Replace('\\', '/'), entry);
        }
            
        templateStream.Position = 0;
        using var originalArchive = new ZipArchive(templateStream, ZipArchiveMode.Read);
        // Create a new zip file for writing

        // Iterate through each entry in the original archive
        foreach (var entry in originalArchive.Entries)
        {
            if (entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) ||
                entry.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) ||
                entry.FullName.Contains("xl/calcChain.xml")
               )
                continue;
                    
            // Create a new entry in the new archive with the same name
            var newEntry = outputFileArchive.ZipFile.CreateEntry(entry.FullName);

            // Copy the content of the original entry to the new entry
            using var originalEntryStream = entry.Open();
            using var newEntryStream = newEntry.Open();
            
            await originalEntryStream.CopyToAsync(newEntryStream
#if NETCOREAPP2_1_OR_GREATER
                    , cancellationToken
#endif
            ).ConfigureAwait(false);
        }

        //read sharedString
        var templateSharedStrings = templateReader.SharedStrings;
        templateStream.Position = 0;

        //read all xlsx sheets
        var templateSheets = templateReader.Archive.ZipFile.Entries
            .Where(w =>
                w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) ||
                w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase));

        int sheetIdx = 0;
        foreach (var templateSheet in templateSheets)
        {
            //every time need to use new XRowInfos or it'll cause duplicate problem: https://user-images.githubusercontent.com/12729184/115003101-0fcab700-9ed8-11eb-9151-ca4d7b86d59e.png
            _xRowInfos = [];
            _xMergeCellInfos = [];
            _newXMergeCellInfos = [];

            var templateSheetStream = templateSheet.Open();
            var templateFullName = templateSheet.FullName;

            var inputValues = _inputValueExtractor.ToValueDictionary(value);
            var outputZipEntry = outputFileArchive.ZipFile.CreateEntry(templateFullName);
            
            using var outputZipSheetEntryStream = outputZipEntry.Open();
            GenerateSheetXmlImplByCreateMode(templateSheet, outputZipSheetEntryStream, templateSheetStream, inputValues, templateSharedStrings, false);
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
            using var calcChainStream = calcChainEntry.Open();
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
                    using var originalEntryStream = entry.Open();
                    using var newEntryStream = newEntry.Open();
                    
                    await originalEntryStream.CopyToAsync(newEntryStream
#if NETCOREAPP2_1_OR_GREATER
                        , cancellationToken
#endif
                    ).ConfigureAwait(false);
                }
            }
        }

        outputFileArchive.ZipFile.Dispose();
    }
}