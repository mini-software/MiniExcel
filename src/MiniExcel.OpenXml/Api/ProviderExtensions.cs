using MiniExcelLib.Core;

// ReSharper disable once CheckNamespace
namespace MiniExcelLib.OpenXml;

public static class ProviderExtensions
{
    public static OpenXmlExporter GetOpenXmlExporter(this MiniExcelExporterProvider exporterProvider) => new(); 
    public static OpenXmlImporter GetOpenXmlImporter(this MiniExcelImporterProvider importerProvider) => new();
    public static OpenXmlTemplater GetOpenXmlTemplater(this MiniExcelTemplaterProvider templaterProvider) => new();
}