namespace MiniExcelLib;

public static class MiniExcel
{
    public static MiniExcelExporter GetExporter() => new();
    public static MiniExcelImporter GetImporter() => new();
    public static MiniExcelTemplater GetTemplater() => new();
}