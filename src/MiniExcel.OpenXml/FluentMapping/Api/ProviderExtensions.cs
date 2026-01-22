using MiniExcelLib.Core;

namespace MiniExcelLib.OpenXml.FluentMapping.Api;

public static class ProviderExtensions
{
    public static MappingExporter GetMappingExporter(this MiniExcelExporterProvider exporterProvider) => new();
    public static MappingExporter GetMappingExporter(this MiniExcelExporterProvider exporterProvider, MappingRegistry registry) => new(registry);

    public static MappingImporter GetMappingImporter(this MiniExcelImporterProvider importerProvider) => new();
    public static MappingImporter GetMappingImporter(this MiniExcelImporterProvider importerProvider, MappingRegistry registry) => new(registry);
    
    public static MappingTemplater GetMappingTemplater(this MiniExcelTemplaterProvider templaterProvider) =>  new();
    public static MappingTemplater GetMappingTemplater(this MiniExcelTemplaterProvider templaterProvider, MappingRegistry registry) =>  new(registry);
}