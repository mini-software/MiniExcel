using MiniExcelLib.Core;

namespace MiniExcelLib.OpenXml.Api;

public static class ProviderExtensions
{
    public static OpenXmlExporter GetOpenXmlExporter(this MiniExcelExporterProvider exporterProvider) => new(); 
    public static OpenXmlImporter GetOpenXmlImporter(this MiniExcelImporterProvider importerProvider) => new();
    public static OpenXmlTemplater GetOpenXmlTemplater(this MiniExcelTemplaterProvider templaterProvider) => new();
}