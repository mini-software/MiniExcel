// ReSharper disable once CheckNamespace
namespace MiniExcelLib.OpenXml;

public static class ProviderExtensions
{
    public static OpenXmlExporter GetOpenXmlExporter(this MiniExcelExporterProvider exporterProvider) => new(); 
    public static OpenXmlImporter GetOpenXmlImporter(this MiniExcelImporterProvider importerProvider) => new();
    public static OpenXmlTemplater GetOpenXmlTemplater(this MiniExcelTemplaterProvider templaterProvider) => new();

    /// <summary>Creates an editor for an existing OpenXml workbook.</summary>
    public static OpenXmlEditor GetOpenXmlEditor(this MiniExcelEditorProvider editorProvider, string path) => new(path);

    /// <summary>Creates an editor for an existing OpenXml workbook stream.</summary>
    public static OpenXmlEditor GetOpenXmlEditor(this MiniExcelEditorProvider editorProvider, Stream stream) => new(stream);
}