namespace MiniExcelLib;

public static class MiniExcel
{
    public static readonly MiniExcelExporterProvider Exporter = new();
    public static readonly MiniExcelImporterProvider Importer = new();
    public static readonly MiniExcelTemplaterProvider Templater = new();
}