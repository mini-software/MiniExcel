using MiniExcelLib.Core;

// ReSharper disable once CheckNamespace
namespace MiniExcelLib;

public static class MiniExcelV2
{
    public static readonly MiniExcelExporterProvider Exporters = new();
    public static readonly MiniExcelImporterProvider Importers = new();
    public static readonly MiniExcelTemplaterProvider Templaters = new();
}

[Obsolete("This class will be removed in the full release, use MiniExcelV2 instead.", true)]
public static class MiniExcel
{
    public static readonly MiniExcelExporterProvider Exporters = new();
    public static readonly MiniExcelImporterProvider Importers = new();
    public static readonly MiniExcelTemplaterProvider Templaters = new();
}
