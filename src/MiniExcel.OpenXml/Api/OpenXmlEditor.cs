using MiniExcelLib.OpenXml.Editor;

// ReSharper disable once CheckNamespace
namespace MiniExcelLib.OpenXml;

public sealed class OpenXmlEditor
{
    internal OpenXmlEditor() { }

    /// <summary>
    /// Creates a new editing pipeline for the provided Excel document.
    /// </summary>
    /// <param name="path">The file path to the Excel document to edit.</param>
    /// <returns>
    /// An <see cref="OpenXmlEditingPipeline"/> instance that can be used to apply modifications to the document.
    /// </returns>
    /// <remarks>
    /// This method opens the file for exclusive read-write access.The file is locked until 
    /// <see cref="OpenXmlEditingPipeline.SaveChangesAsync"/> is called.
    /// </remarks>
    public OpenXmlEditingPipeline StartEditingPipeline(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path cannot be null or whitespace.", nameof(path));

        var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
        return new OpenXmlEditingPipeline(stream, leaveOpen: false);
    }

    /// <summary>
    /// Creates a new editing pipeline for the provided Excel document.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data.</param>
    /// <param name="leaveOpen">
    /// If true the stream remains open after changes are saved and must be disposed by the caller,
    /// if false it is automatically closed when changes are saved. Default is false.
    /// </param>
    /// <returns>
    /// An <see cref="OpenXmlEditingPipeline"/> instance that can be used to apply modifications to the file.
    /// </returns>
    /// <remarks>
    /// Even with parameter <c>leaveOpen: false</c>, the underlying stream will not be disposed until
    /// <see cref="OpenXmlEditingPipeline.SaveChanges"/> or <see cref="OpenXmlEditingPipeline.SaveChangesAsync"/> are called.
    /// </remarks>
    public OpenXmlEditingPipeline StartEditingPipeline(Stream stream, bool leaveOpen = false)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

        return new OpenXmlEditingPipeline(stream, leaveOpen);
    }
}
