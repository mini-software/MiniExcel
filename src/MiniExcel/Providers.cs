namespace MiniExcelLib;

public sealed class MiniExcelImporterProvider
{
    public OpenXmlImporter GetExcelImporter() => new();
}

public sealed class MiniExcelExporterProvider
{
    public OpenXmlExporter GetExcelExporter() => new();
}

public sealed class MiniExcelTemplaterProvider
{
    public OpenXmlTemplater GetExcelTemplater() => new();
}
