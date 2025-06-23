using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using MiniExcelLibs.Csv;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.SaveByTemplate;
using Zomp.SyncMethodGenerator;
using ExcelOpenXmlTemplate = MiniExcelLibs.SaveByTemplate.ExcelOpenXmlTemplate;

namespace MiniExcelLibs;

internal static partial class ExcelReaderFactory
{
    [CreateSyncVersion]
    internal static async Task<IExcelReader> GetProviderAsync(Stream stream, ExcelType excelType, IMiniExcelConfiguration? configuration, CancellationToken cancellationToken = default)
    {
        return excelType switch
        {
            ExcelType.CSV => new CsvReader(stream, configuration),
            ExcelType.XLSX => await ExcelOpenXmlSheetReader.CreateAsync(stream, configuration, cancellationToken: cancellationToken).ConfigureAwait(false),
            _ => throw new NotSupportedException("Something went wrong. Please report this issue you are experiencing with MiniExcel.")
        };
    }
}

internal static class ExcelWriterFactory
{
    internal static IExcelWriter GetProvider(Stream stream, object value, string? sheetName, ExcelType excelType, IMiniExcelConfiguration? configuration, bool printHeader)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("Sheet names cannot be empty or null", nameof(sheetName));
        if (sheetName?.Length > 31 && excelType == ExcelType.XLSX)
            throw new ArgumentException("Sheet names must be less than 31 characters", nameof(sheetName));
        if (excelType == ExcelType.UNKNOWN)
            throw new ArgumentException("Excel type cannot be ExcelType.UNKNOWN", nameof(excelType));

        return excelType switch
        {
            ExcelType.CSV => new CsvWriter(stream, value, configuration, printHeader),
            ExcelType.XLSX => new ExcelOpenXmlSheetWriter(stream, value, sheetName, configuration, printHeader),
            _ => throw new NotSupportedException($"The {excelType} Excel format is not supported")
        };
    }
}

internal static class ExcelTemplateFactory
{
    internal static IExcelTemplate GetProvider(Stream stream, IMiniExcelConfiguration? configuration, ExcelType excelType = ExcelType.XLSX)
    {
        if (excelType != ExcelType.XLSX)
            throw new NotSupportedException("Something went wrong. Please report this issue you are experiencing with MiniExcel.");

        var valueExtractor = new InputValueExtractor();
        return new ExcelOpenXmlTemplate(stream, configuration, valueExtractor);
    }
}