namespace MiniExcelLib.Csv;

public static class ProviderExtensions
{
    public static Api.CsvExporter GetCsvExporter(this MiniExcelExporterProvider exporterProvider) => new(); 
    public static Api.CsvImporter GetCsvImporter(this MiniExcelImporterProvider importerProvider) => new();
}