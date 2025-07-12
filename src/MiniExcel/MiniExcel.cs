namespace MiniExcelLib;

public static class MiniExcel
{
    public static MiniExcelExporterProvider GetExporterProvider() => new();
    public static MiniExcelImporterProvider GetImporterProvider() => new();
    public static MiniExcelTemplaterProvider GetTemplaterProvider() => new();
}