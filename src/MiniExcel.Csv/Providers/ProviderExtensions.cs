namespace MiniExcelLib.Csv.Providers;

public static class ProviderExtensions
{
    public static CsvExporter GetCsvExporter(this MiniExcelExporterProvider exporterProvider) => new(); 
    public static CsvImporter GetCsvImporter(this MiniExcelImporterProvider importerProvider) => new();
}