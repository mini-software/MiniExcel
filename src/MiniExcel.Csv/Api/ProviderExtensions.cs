namespace MiniExcelLib.Csv.Api;

public static class ProviderExtensions
{
    public static CsvExporter GetCsvExporter(this MiniExcelExporterProvider exporterProvider) => new(); 
    public static CsvImporter GetCsvImporter(this MiniExcelImporterProvider importerProvider) => new();
}