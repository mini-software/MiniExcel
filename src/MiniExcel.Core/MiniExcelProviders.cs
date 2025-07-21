namespace MiniExcelLib.Core;

public sealed class MiniExcelImporterProvider
{
    internal MiniExcelImporterProvider() { }
    
    public OpenXmlImporter GetOpenXmlImporter() => new();
}

public sealed class MiniExcelExporterProvider
{
    internal MiniExcelExporterProvider() { }

    public OpenXmlExporter GetOpenXmlExporter() => new();
}

public sealed class MiniExcelTemplaterProvider
{
    internal MiniExcelTemplaterProvider() { }

    public OpenXmlTemplater GetOpenXmlTemplater() => new();
}
