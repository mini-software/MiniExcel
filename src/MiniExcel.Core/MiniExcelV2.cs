using MiniExcelLib.Core;

// ReSharper disable once CheckNamespace
namespace MiniExcelLib;

public static class MiniExcelV2
{
    public static readonly MiniExcelExporterProvider Exporters = new();
    public static readonly MiniExcelImporterProvider Importers = new();
    public static readonly MiniExcelTemplaterProvider Templaters = new();
    public static readonly MiniExcelEditorProvider Editors = new();
}
