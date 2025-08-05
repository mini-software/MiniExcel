namespace MiniExcelLib.Core;

public sealed class MiniExcelImporterProvider
{
    internal MiniExcelImporterProvider() { }
    
    public OpenXmlImporter GetOpenXmlImporter() => new();
    public MappingImporter GetMappingImporter() => new();
    public MappingImporter GetMappingImporter(MappingRegistry registry) => new(registry);
}

public sealed class MiniExcelExporterProvider
{
    internal MiniExcelExporterProvider() { }

    public OpenXmlExporter GetOpenXmlExporter() => new();
    public MappingExporter GetMappingExporter() => new();
    public MappingExporter GetMappingExporter(MappingRegistry registry) => new(registry);
}

public sealed class MiniExcelTemplaterProvider
{
    internal MiniExcelTemplaterProvider() { }

    public OpenXmlTemplater GetOpenXmlTemplater() => new();
}
