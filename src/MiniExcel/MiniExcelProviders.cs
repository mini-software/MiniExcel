namespace MiniExcelLib;

public sealed class MiniExcelImporterProvider
{
    internal MiniExcelImporterProvider() { }
    
    public OpenXmlImporter GetExcelImporter() => new();
}

public sealed class MiniExcelExporterProvider
{
    internal MiniExcelExporterProvider() { }

    public OpenXmlExporter GetExcelExporter() => new();
}

public sealed class MiniExcelTemplaterProvider
{
    internal MiniExcelTemplaterProvider() { }

    public OpenXmlTemplater GetExcelTemplater() => new();
}
