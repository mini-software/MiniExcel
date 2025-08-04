using MiniExcelLib.Core;

// ReSharper disable once CheckNamespace
namespace MiniExcelLib.Csv;

public static class ProviderExtensions
{
    public static CsvExporter GetCsvExporter(this MiniExcelExporterProvider exporterProvider) => new(); 
    public static CsvImporter GetCsvImporter(this MiniExcelImporterProvider importerProvider) => new();
}