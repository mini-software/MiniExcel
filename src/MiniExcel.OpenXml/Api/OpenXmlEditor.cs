using MiniExcelLib.OpenXml.Editor;

// ReSharper disable once CheckNamespace
namespace MiniExcelLib.OpenXml;

public sealed class OpenXmlEditor
{
    internal OpenXmlEditor() { }

    
    public OpenXmlEditingPipeline StartEditingPipeline(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path cannot be null or whitespace.", nameof(path));

        var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
        return new OpenXmlEditingPipeline(stream, leaveOpen: false);
    }

    public OpenXmlEditingPipeline StartEditingPipeline(Stream stream, bool leaveOpen = false)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

        return new OpenXmlEditingPipeline(stream, leaveOpen);
    }
}
