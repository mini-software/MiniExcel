using System.Data;
using MiniExcelLib.Core.DataReader;
using MiniExcelLib.Core.OpenXml.Models;
using MiniExcelLib.Core.OpenXml.Picture;
using MiniExcelLib.Csv;
using MiniExcelLibs.OpenXml;
using Zomp.SyncMethodGenerator;

using NewMiniExcel = MiniExcelLib.Core.MiniExcel;
using OpenXmlExporter = MiniExcelLib.Core.OpenXmlExporter;
using OpenXmlImporter = MiniExcelLib.Core.OpenXmlImporter;
using OpenXmlTemplater = MiniExcelLib.Core.OpenXmlTemplater;

// ReSharper disable once CheckNamespace
namespace MiniExcelLibs;

public static partial class MiniExcel
{
    private static readonly OpenXmlExporter ExcelExporter = NewMiniExcel.Exporters.GetOpenXmlExporter();
    private static readonly OpenXmlImporter ExcelImporter = NewMiniExcel.Importers.GetOpenXmlImporter();
    private static readonly OpenXmlTemplater ExcelTemplater = NewMiniExcel.Templaters.GetOpenXmlTemplater();
    
    private static readonly CsvExporter CsvExporter = NewMiniExcel.Exporters.GetCsvExporter();
    private static readonly CsvImporter CsvImporter = NewMiniExcel.Importers.GetCsvImporter();

    
    [CreateSyncVersion]
    public static async Task AddPictureAsync(string path, CancellationToken cancellationToken = default, params MiniExcelPicture[] images) 
        => await ExcelTemplater.AddPictureAsync(path, cancellationToken, images).ConfigureAwait(false);
    
    [CreateSyncVersion]
    public static async Task AddPictureAsync(Stream excelStream, CancellationToken cancellationToken = default, params MiniExcelPicture[] images) 
        =>  await ExcelTemplater.AddPictureAsync(excelStream, cancellationToken, images).ConfigureAwait(false);
    
    public static MiniExcelDataReader GetReader(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.GetDataReader(path, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration),
            ExcelType.CSV => CsvImporter.GetDataReader(path, useHeaderRow, configuration as Csv.CsvConfiguration),
            _ => throw new NotSupportedException($"Excel type {type} is not a valid Excel type")
        };
    }
    
    public static MiniExcelDataReader GetReader(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.GetDataReader(stream, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration),
            ExcelType.CSV => CsvImporter.GetDataReader(stream, useHeaderRow, configuration as Csv.CsvConfiguration),
            _ => throw new NotSupportedException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    public static async Task<int> InsertAsync(string path, object value, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration? configuration = null, bool printHeader = true, bool overwriteSheet = false, CancellationToken cancellationToken = default)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelExporter.InsertSheetAsync(path, value, sheetName, printHeader, overwriteSheet, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvExporter.AppendAsync(path, value, printHeader, configuration as Csv.CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    public static async Task<int> InsertAsync(this Stream stream, object value, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration? configuration = null, bool printHeader = true, bool overwriteSheet = false, CancellationToken cancellationToken = default)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelExporter.InsertSheetAsync(stream, value, sheetName, printHeader, overwriteSheet, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvExporter.AppendAsync(stream, value, configuration as Csv.CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    public static async Task<int[]> SaveAsAsync(string path, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration? configuration = null, bool overwriteFile = false, CancellationToken cancellationToken = default, IProgress<int>? progress = null)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelExporter.ExportAsync(path, value, printHeader, sheetName, printHeader, configuration as OpenXmlConfiguration, cancellationToken, progress).ConfigureAwait(false),
            ExcelType.CSV => await CsvExporter.ExportAsync(path, value, printHeader, overwriteFile, configuration as Csv.CsvConfiguration, cancellationToken, progress).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    public static async Task<int[]> SaveAsAsync(this Stream stream, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration? configuration = null, CancellationToken  cancellationToken = default, IProgress<int>? progress = null)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelExporter.ExportAsync(stream, value, printHeader, sheetName, configuration as OpenXmlConfiguration, cancellationToken, progress).ConfigureAwait(false),
            ExcelType.CSV => await CsvExporter.ExportAsync(stream, value, printHeader, configuration as Csv.CsvConfiguration, cancellationToken, progress).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    public static IAsyncEnumerable<T> QueryAsync<T>(string path, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, bool hasHeader = true, CancellationToken cancellationToken = default) where T : class, new()
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryAsync<T>(path, sheetName, startCell, hasHeader, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => CsvImporter.QueryAsync<T>(path, hasHeader, configuration as Csv.CsvConfiguration, cancellationToken),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    public static IAsyncEnumerable<T> QueryAsync<T>(this Stream stream, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, bool hasHeader = true, CancellationToken cancellationToken = default) where T : class, new()
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryAsync<T>(stream, sheetName, startCell, hasHeader, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => CsvImporter.QueryAsync<T>(stream, hasHeader, configuration as Csv.CsvConfiguration, cancellationToken),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }
   
    [CreateSyncVersion]
    public static IAsyncEnumerable<dynamic> QueryAsync(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryAsync(path, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => CsvImporter.QueryAsync(path, useHeaderRow, configuration as Csv.CsvConfiguration, cancellationToken),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }
    
    [CreateSyncVersion]
    public static IAsyncEnumerable<dynamic> QueryAsync(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryAsync(stream, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => CsvImporter.QueryAsync(stream, useHeaderRow, configuration as Csv.CsvConfiguration, cancellationToken),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    #region QueryRange

    /// <summary>
    /// Extract the given range。 Only uppercase letters are effective。
    /// e.g.
    ///     MiniExcel.QueryRange(path, startCell: "A2", endCell: "C3")
    ///     A2 represents the second row of column A, C3 represents the third row of column C
    ///     If you don't want to restrict rows, just don't include numbers
    /// </summary>
    [CreateSyncVersion]
    public static IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", string endCell = "", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryRangeAsync(path, useHeaderRow, sheetName, startCell, endCell, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => throw new NotSupportedException("QueryRange is not supported for csv"),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    public static IAsyncEnumerable<dynamic> QueryRangeAsync(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", string endCell = "", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryRangeAsync(stream, useHeaderRow, sheetName, startCell, endCell, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => CsvImporter.QueryAsync(stream, useHeaderRow, configuration as Csv.CsvConfiguration, cancellationToken),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    public static IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null, int? endColumnIndex = null, IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryRangeAsync(path, useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => throw new NotSupportedException("QueryRange is not supported for csv"),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }
    
    [CreateSyncVersion]
    public static IAsyncEnumerable<dynamic> QueryRangeAsync(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null, int? endColumnIndex = null, IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryRangeAsync(stream, useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => throw new NotSupportedException("QueryRange is not supported for csv"),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    #endregion QueryRange

    [CreateSyncVersion]
    public static async Task SaveAsByTemplateAsync(string path, string templatePath, object value, IConfiguration? configuration = null, CancellationToken cancellationToken = default) 
        => await ExcelTemplater.ApplyTemplateAsync(path, templatePath, value, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false);

    [CreateSyncVersion]
    public static async Task SaveAsByTemplateAsync(string path, byte[] templateBytes, object value, IConfiguration? configuration = null)
        => await ExcelTemplater.ApplyTemplateAsync(path, templateBytes, value, configuration as OpenXmlConfiguration).ConfigureAwait(false);
    
    [CreateSyncVersion]
    public static async Task SaveAsByTemplateAsync(this Stream stream, string templatePath, object value, IConfiguration? configuration = null)
        => await ExcelTemplater.ApplyTemplateAsync(stream, templatePath, value, configuration as OpenXmlConfiguration).ConfigureAwait(false);

    [CreateSyncVersion]
    public static async Task SaveAsByTemplateAsync(this Stream stream, byte[] templateBytes, object value, IConfiguration? configuration = null)
        => await ExcelTemplater.ApplyTemplateAsync(stream, templateBytes, value, configuration as OpenXmlConfiguration).ConfigureAwait(false);
    
    [CreateSyncVersion]
    public static async Task SaveAsByTemplateAsync(string path, Stream templateStream, object value, IConfiguration? configuration = null)
        => await ExcelTemplater.ApplyTemplateAsync(path, templateStream, value, configuration as OpenXmlConfiguration).ConfigureAwait(false);

    [CreateSyncVersion]
    public static async Task SaveAsByTemplateAsync(this Stream stream, Stream templateStream, object value, IConfiguration? configuration = null)
        => await ExcelTemplater.ApplyTemplateAsync(stream, templateStream, value, configuration as OpenXmlConfiguration).ConfigureAwait(false);

    #region MergeCells

    [CreateSyncVersion]
    public static async Task MergeSameCellsAsync(string mergedFilePath, string path, ExcelType excelType = ExcelType.XLSX, IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        if (excelType != ExcelType.XLSX) 
            throw new NotSupportedException("MergeSameCells is only supported for Xlsx files");

        await ExcelTemplater.MergeSameCellsAsync(mergedFilePath, path, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task MergeSameCellsAsync(this Stream stream, string path, ExcelType excelType = ExcelType.XLSX, IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        if (excelType != ExcelType.XLSX) 
            throw new NotSupportedException("MergeSameCells is only supported for Xlsx files");

        await ExcelTemplater.MergeSameCellsAsync(stream, path, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task MergeSameCellsAsync(this Stream stream, byte[] filePath, ExcelType excelType = ExcelType.XLSX, IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        if (excelType != ExcelType.XLSX) 
            throw new NotSupportedException("MergeSameCells is only supported for Xlsx files");

        await ExcelTemplater.MergeSameCellsAsync(stream, filePath, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false);
    }

    #endregion

    [CreateSyncVersion]
    [Obsolete("The use of QueryAsDataTable is not recommended, because it'll load all data into memory.")]
    public static async Task<DataTable> QueryAsDataTableAsync(string path, bool useHeaderRow = true, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelImporter.QueryAsDataTableAsync(path, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvImporter.QueryAsDataTableAsync(path, useHeaderRow, configuration as Csv.CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    [Obsolete("The use of QueryAsDataTable is not recommended, because it'll load all data into memory.")]
    public static async Task<DataTable> QueryAsDataTableAsync(this Stream stream, bool useHeaderRow = true, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelImporter.QueryAsDataTableAsync(stream, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvImporter.QueryAsDataTableAsync(stream, useHeaderRow, configuration as Csv.CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    public static async Task<List<string>> GetSheetNamesAsync(string path, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
        => await ExcelImporter.GetSheetNamesAsync(path, config, cancellationToken).ConfigureAwait(false);
    
    [CreateSyncVersion]
    public static async Task<List<string>> GetSheetNamesAsync(this Stream stream, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
        => await ExcelImporter.GetSheetNamesAsync(stream, config, cancellationToken).ConfigureAwait(false);

    [CreateSyncVersion]
    public static async Task<List<SheetInfo>> GetSheetInformationsAsync(string path, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
        => await ExcelImporter.GetSheetInformationsAsync(path, config, cancellationToken).ConfigureAwait(false);
    
    [CreateSyncVersion]
    public static async Task<List<SheetInfo>> GetSheetInformationsAsync(this Stream stream, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
        => await ExcelImporter.GetSheetInformationsAsync(stream, config, cancellationToken).ConfigureAwait(false);

    [CreateSyncVersion]
    public static async Task<ICollection<string>> GetColumnsAsync(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelImporter.GetColumnNamesAsync(path, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvImporter.GetColumnNamesAsync(path, useHeaderRow, configuration as Csv.CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    public static async Task<ICollection<string>> GetColumnsAsync(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelImporter.GetColumnNamesAsync(stream, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvImporter.GetColumnNamesAsync(stream, useHeaderRow, configuration as Csv.CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    public static async Task<IList<ExcelRange>> GetSheetDimensionsAsync(string path, CancellationToken cancellationToken = default) 
        => await ExcelImporter.GetSheetDimensionsAsync(path, cancellationToken).ConfigureAwait(false);
    
    [CreateSyncVersion]
    public static async Task<IList<ExcelRange>> GetSheetDimensionsAsync(this Stream stream, CancellationToken cancellationToken = default) 
        => await ExcelImporter.GetSheetDimensionsAsync(stream, cancellationToken).ConfigureAwait(false);

    [CreateSyncVersion]
    public static async Task ConvertCsvToXlsxAsync(string csv, string xlsx, CancellationToken cancellationToken = default)
        => await CsvExporter.ConvertCsvToXlsxAsync(csv, xlsx, cancellationToken: cancellationToken).ConfigureAwait(false);
    
    [CreateSyncVersion]
    public static async Task ConvertCsvToXlsxAsync(Stream csv, Stream xlsx, CancellationToken cancellationToken = default)
        => await CsvExporter.ConvertCsvToXlsxAsync(csv, xlsx, cancellationToken: cancellationToken).ConfigureAwait(false);

    [CreateSyncVersion]
    public static async Task ConvertXlsxToCsvAsync(string xlsx, string csv, CancellationToken cancellationToken = default)
        => await CsvExporter.ConvertXlsxToCsvAsync(xlsx, csv, cancellationToken: cancellationToken).ConfigureAwait(false);
    
    [CreateSyncVersion]
    public static async Task ConvertXlsxToCsvAsync(Stream xlsx, Stream csv, CancellationToken cancellationToken = default)
        => await CsvExporter.ConvertXlsxToCsvAsync(xlsx, csv, cancellationToken: cancellationToken).ConfigureAwait(false);
}