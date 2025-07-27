namespace MiniExcelLib.Core;

public static class MiniExcel
{
    public static readonly MiniExcelExporterProvider Exporters = new();
    public static readonly MiniExcelImporterProvider Importers = new();
    public static readonly MiniExcelTemplaterProvider Templaters = new();
}