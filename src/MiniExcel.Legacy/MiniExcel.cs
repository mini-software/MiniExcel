using System.Data;
using MiniExcelLib;
using MiniExcelLib.Csv;
using MiniExcelLib.DataReader;
using MiniExcelLib.OpenXml.Models;
using MiniExcelLib.OpenXml.Picture;
using Zomp.SyncMethodGenerator;
using CsvExporter = MiniExcelLib.Csv.Api.CsvExporter;
using CsvImporter = MiniExcelLib.Csv.Api.CsvImporter;
using MiniExcelNew = MiniExcelLib.MiniExcel;
using OpenXmlExporter = MiniExcelLib.OpenXml.Api.OpenXmlExporter;
using OpenXmlImporter = MiniExcelLib.OpenXml.Api.OpenXmlImporter;
using OpenXmlTemplater = MiniExcelLib.OpenXml.Api.OpenXmlTemplater;

namespace MiniExcelLibs;

public static partial class MiniExcel
{
    private static readonly OpenXmlExporter ExcelExporter = MiniExcelNew.Exporter.GetExcelExporter();
    private static readonly OpenXmlImporter ExcelImporter = MiniExcelNew.Importer.GetExcelImporter();
    private static readonly OpenXmlTemplater ExcelTemplater = MiniExcelNew.Templater.GetExcelTemplater();
    
    private static readonly CsvExporter CsvExporter = MiniExcelNew.Exporter.GetCsvExporter();
    private static readonly CsvImporter CsvImporter = MiniExcelNew.Importer.GetCsvImporter();

    
    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Exporter instead.")]
    public static async Task AddPictureAsync(string path, CancellationToken cancellationToken = default, params MiniExcelPicture[] images) 
        => await ExcelExporter.AddExcelPictureAsync(path, cancellationToken, images).ConfigureAwait(false);
    
    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Exporter instead.")]
    public static async Task AddPictureAsync(Stream excelStream, CancellationToken cancellationToken = default, params MiniExcelPicture[] images) 
        =>  await ExcelExporter.AddExcelPictureAsync(excelStream, cancellationToken, images).ConfigureAwait(false);

    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static MiniExcelDataReader GetReader(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.GetExcelDataReader(path, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration),
            ExcelType.CSV => CsvImporter.GetCsvDataReader(path, useHeaderRow, configuration as CsvConfiguration),
            _ => throw new NotSupportedException($"Excel type {type} is not a valid Excel type")
        };
    }

    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static MiniExcelDataReader GetReader(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.GetExcelDataReader(stream, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration),
            ExcelType.CSV => CsvImporter.GetCsvDataReader(stream, useHeaderRow, configuration as CsvConfiguration),
            _ => throw new NotSupportedException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Exporter instead.")]
    public static async Task<int> InsertAsync(string path, object value, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration? configuration = null, bool printHeader = true, bool overwriteSheet = false, CancellationToken cancellationToken = default)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelExporter.InsertExcelSheetAsync(path, value, sheetName, printHeader, overwriteSheet, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvExporter.AppendToCsvAsync(path, value, printHeader, configuration as CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Exporter instead.")]
    public static async Task<int> InsertAsync(this Stream stream, object value, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration? configuration = null, bool printHeader = true, bool overwriteSheet = false, CancellationToken cancellationToken = default)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelExporter.InsertExcelSheetAsync(stream, value, sheetName, printHeader, overwriteSheet, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvExporter.AppendToCsvAsync(stream, value, configuration as CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Exporter instead.")]
    public static async Task<int[]> SaveAsAsync(string path, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration? configuration = null, bool overwriteFile = false, CancellationToken cancellationToken = default)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelExporter.ExportExcelAsync(path, value, printHeader, sheetName, printHeader, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvExporter.ExportCsvAsync(path, value, printHeader, overwriteFile, configuration as CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Exporter instead.")]
    public static async Task<int[]> SaveAsAsync(this Stream stream, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration? configuration = null, CancellationToken  cancellationToken = default)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelExporter.ExportExcelAsync(stream, value, printHeader, sheetName, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvExporter.ExportCsvAsync(stream, value, printHeader, configuration as CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static IAsyncEnumerable<T> QueryAsync<T>(string path, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, bool hasHeader = true, CancellationToken cancellationToken = default) where T : class, new()
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryExcelAsync<T>(path, sheetName, startCell, hasHeader, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => CsvImporter.QueryCsvAsync<T>(path, hasHeader, configuration as CsvConfiguration, cancellationToken),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static IAsyncEnumerable<T> QueryAsync<T>(this Stream stream, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, bool hasHeader = true, CancellationToken cancellationToken = default) where T : class, new()
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryExcelAsync<T>(stream, sheetName, startCell, hasHeader, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => CsvImporter.QueryCsvAsync<T>(stream, hasHeader, configuration as CsvConfiguration, cancellationToken),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }
   
    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static IAsyncEnumerable<dynamic> QueryAsync(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryExcelAsync(path, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => CsvImporter.QueryCsvAsync(path, useHeaderRow, configuration as CsvConfiguration, cancellationToken),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }
    
    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static IAsyncEnumerable<dynamic> QueryAsync(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryExcelAsync(stream, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => CsvImporter.QueryCsvAsync(stream, useHeaderRow, configuration as CsvConfiguration, cancellationToken),
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
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", string endCell = "", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryExcelRangeAsync(path, useHeaderRow, sheetName, startCell, endCell, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => throw new NotSupportedException("QueryRange is not supported for csv"),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static IAsyncEnumerable<dynamic> QueryRangeAsync(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", string endCell = "", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryExcelRangeAsync(stream, useHeaderRow, sheetName, startCell, endCell, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => CsvImporter.QueryCsvAsync(stream, useHeaderRow, configuration as CsvConfiguration, cancellationToken),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null, int? endColumnIndex = null, IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryExcelRangeAsync(path, useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => throw new NotSupportedException("QueryRange is not supported for csv"),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }
    
    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static IAsyncEnumerable<dynamic> QueryRangeAsync(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null, int? endColumnIndex = null, IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => ExcelImporter.QueryExcelRangeAsync(stream, useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, configuration as OpenXmlConfiguration, cancellationToken),
            ExcelType.CSV => throw new NotSupportedException("QueryRange is not supported for csv"),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    #endregion QueryRange

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Exporter instead.")]
    public static async Task SaveAsByTemplateAsync(string path, string templatePath, object value, IConfiguration? configuration = null, CancellationToken cancellationToken = default) 
        => await ExcelTemplater.ApplyXlsxTemplateAsync(path, templatePath, value, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false);

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Templater instead.")]
    public static async Task SaveAsByTemplateAsync(string path, byte[] templateBytes, object value, IConfiguration? configuration = null)
        => await ExcelTemplater.ApplyXlsxTemplateAsync(path, templateBytes, value, configuration as OpenXmlConfiguration).ConfigureAwait(false);
    
    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Templater instead.")]
    public static async Task SaveAsByTemplateAsync(this Stream stream, string templatePath, object value, IConfiguration? configuration = null)
        => await ExcelTemplater.ApplyXlsxTemplateAsync(stream, templatePath, value, configuration as OpenXmlConfiguration).ConfigureAwait(false);

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Templater instead.")]
    public static async Task SaveAsByTemplateAsync(this Stream stream, byte[] templateBytes, object value, IConfiguration? configuration = null)
        => await ExcelTemplater.ApplyXlsxTemplateAsync(stream, templateBytes, value, configuration as OpenXmlConfiguration).ConfigureAwait(false);
    
    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Templater instead.")]
    public static async Task SaveAsByTemplateAsync(string path, Stream templateStream, object value, IConfiguration? configuration = null)
        => await ExcelTemplater.ApplyXlsxTemplateAsync(path, templateStream, value, configuration as OpenXmlConfiguration).ConfigureAwait(false);

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Templater instead.")]
    public static async Task SaveAsByTemplateAsync(this Stream stream, Stream templateStream, object value, IConfiguration? configuration = null)
        => await ExcelTemplater.ApplyXlsxTemplateAsync(stream, templateStream, value, configuration as OpenXmlConfiguration).ConfigureAwait(false);

    #region MergeCells

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Templater instead.")]
    public static async Task MergeSameCellsAsync(string mergedFilePath, string path, ExcelType excelType = ExcelType.XLSX, IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        if (excelType != ExcelType.XLSX) 
            throw new NotSupportedException("MergeSameCells is only supported for Xlsx files");

        await ExcelTemplater.MergeSameCellsAsync(mergedFilePath, path, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Templater instead.")]
    public static async Task MergeSameCellsAsync(this Stream stream, string path, ExcelType excelType = ExcelType.XLSX, IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        if (excelType != ExcelType.XLSX) 
            throw new NotSupportedException("MergeSameCells is only supported for Xlsx files");

        await ExcelTemplater.MergeSameCellsAsync(stream, path, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Templater instead.")]
    public static async Task MergeSameCellsAsync(this Stream stream, byte[] filePath, ExcelType excelType = ExcelType.XLSX, IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        if (excelType != ExcelType.XLSX) 
            throw new NotSupportedException("MergeSameCells is only supported for Xlsx files");

        await ExcelTemplater.MergeSameCellsAsync(stream, filePath, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false);
    }

    #endregion

    /// <summary>
    /// The use of QueryAsDataTable is not recommended, because it'll load all data into memory.
    /// </summary>
    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static async Task<DataTable> QueryAsDataTableAsync(string path, bool useHeaderRow = true, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelImporter.QueryExcelAsDataTableAsync(path, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvImporter.QueryCsvAsDataTableAsync(path, useHeaderRow, configuration as CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    /// <summary>
    /// The use of QueryAsDataTable is not recommended, because it'll load all data into memory.
    /// </summary>
    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static async Task<DataTable> QueryAsDataTableAsync(this Stream stream, bool useHeaderRow = true, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelImporter.QueryExcelAsDataTableAsync(stream, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvImporter.QueryCsvAsDataTableAsync(stream, useHeaderRow, configuration as CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion] 
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static async Task<List<string>> GetSheetNamesAsync(string path, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
        => await ExcelImporter.GetSheetNamesAsync(path, config, cancellationToken).ConfigureAwait(false);
    
    [CreateSyncVersion] 
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static async Task<List<string>> GetSheetNamesAsync(this Stream stream, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
        => await ExcelImporter.GetSheetNamesAsync(stream, config, cancellationToken).ConfigureAwait(false);

    [CreateSyncVersion] 
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static async Task<List<SheetInfo>> GetSheetInformationsAsync(string path, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
        => await ExcelImporter.GetSheetInformationsAsync(path, config, cancellationToken).ConfigureAwait(false);
    
    [CreateSyncVersion] 
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static async Task<List<SheetInfo>> GetSheetInformationsAsync(this Stream stream, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
        => await ExcelImporter.GetSheetInformationsAsync(stream, config, cancellationToken).ConfigureAwait(false);

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static async Task<ICollection<string>> GetColumnsAsync(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = path.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelImporter.GetExcelColumnsAsync(path, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvImporter.GetCsvColumnsAsync(path, useHeaderRow, configuration as CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static async Task<ICollection<string>> GetColumnsAsync(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var type = stream.GetExcelType(excelType);
        return type switch
        {
            ExcelType.XLSX => await ExcelImporter.GetExcelColumnsAsync(stream, useHeaderRow, sheetName, startCell, configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false),
            ExcelType.CSV => await CsvImporter.GetCsvColumnsAsync(stream, useHeaderRow, configuration as CsvConfiguration, cancellationToken).ConfigureAwait(false),
            _ => throw new InvalidDataException($"Excel type {type} is not a valid Excel type")
        };
    }

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static async Task<IList<ExcelRange>> GetSheetDimensionsAsync(string path, CancellationToken cancellationToken = default) 
        => await ExcelImporter.GetSheetDimensionsAsync(path, cancellationToken).ConfigureAwait(false);
    
    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Importer instead.")]
    public static async Task<IList<ExcelRange>> GetSheetDimensionsAsync(this Stream stream, CancellationToken cancellationToken = default) 
        => await ExcelImporter.GetSheetDimensionsAsync(stream, cancellationToken).ConfigureAwait(false);

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Exporter instead.")]
    public static async Task ConvertCsvToXlsxAsync(string csv, string xlsx, CancellationToken cancellationToken = default)
        => await CsvExporter.ConvertCsvToXlsxAsync(csv, xlsx, cancellationToken: cancellationToken).ConfigureAwait(false);
    
    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Exporter instead.")]
    public static async Task ConvertCsvToXlsxAsync(Stream csv, Stream xlsx, CancellationToken cancellationToken = default)
        => await CsvExporter.ConvertCsvToXlsxAsync(csv, xlsx, cancellationToken: cancellationToken).ConfigureAwait(false);

    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Exporter instead.")]
    public static async Task ConvertXlsxToCsvAsync(string xlsx, string csv, CancellationToken cancellationToken = default)
        => await CsvExporter.ConvertXlsxToCsvAsync(xlsx, csv, cancellationToken: cancellationToken).ConfigureAwait(false);
    
    [CreateSyncVersion]
    [Obsolete("This is a legacy method signature that will be removed in a future version. Please use the methods from one of the providers in MiniExcelLib.MiniExcel.Exporter instead.")]
    public static async Task ConvertXlsxToCsvAsync(Stream xlsx, Stream csv, CancellationToken cancellationToken = default)
        => await CsvExporter.ConvertXlsxToCsvAsync(xlsx, csv, cancellationToken: cancellationToken).ConfigureAwait(false);

}