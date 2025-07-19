namespace MiniExcelLib;

public sealed class MiniExcelImporterProvider
{
    internal MiniExcelImporterProvider() { }
    
    public OpenXml.Api.OpenXmlImporter GetExcelImporter() => new();
}

public sealed class MiniExcelExporterProvider
{
    internal MiniExcelExporterProvider() { }

    public OpenXml.Api.OpenXmlExporter GetExcelExporter() => new();
}

public sealed class MiniExcelTemplaterProvider
{
    internal MiniExcelTemplaterProvider() { }

    public OpenXml.Api.OpenXmlTemplater GetExcelTemplater() => new();
}
