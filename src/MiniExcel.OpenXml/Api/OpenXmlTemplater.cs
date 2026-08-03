using MiniExcelLib.OpenXml.Picture;
using MiniExcelLib.OpenXml.Templates;
using OpenXmlTemplate = MiniExcelLib.OpenXml.Templates.OpenXmlTemplate;

// ReSharper disable once CheckNamespace
namespace MiniExcelLib.OpenXml;

public sealed partial class OpenXmlTemplater
{
    internal OpenXmlTemplater() { }
    
    /// <summary>
    /// Adds pictures to an existing OpenXml document.
    /// </summary>
    /// <param name="path">The path to the OpenXml document to modify. The stream must be readable and writable.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <param name="images">A parameter array of <see cref="MiniExcelPicture"/> objects representing the pictures to add to the document.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    [CreateSyncVersion]
    public async Task AddPictureAsync(string path, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
    {
        var stream = File.Open(path, FileMode.OpenOrCreate);
        await using var disposableStream = stream.ConfigureAwait(false); 

        await AddPictureAsync(stream, cancellationToken, images).ConfigureAwait(false);
    }

    /// <summary>
    /// Adds pictures to an existing OpenXml document.
    /// </summary>
    /// <param name="stream">The stream containing the OpenXml document to modify. The stream must be readable and writable.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <param name="images">A parameter array of <see cref="MiniExcelPicture"/> objects representing the pictures to add to the document.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    [CreateSyncVersion]
    public async Task AddPictureAsync(Stream stream, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
    {
        await MiniExcelPictureImplement.AddPictureAsync(stream, cancellationToken, images).ConfigureAwait(false);
    }

    /// <summary>
    /// Fills a template with the provided data and saves the result to a file.
    /// </summary>
    /// <param name="path">The destination file to write the filled document to.</param>
    /// <param name="templatePath">The path to the OpenXml template document.</param>
    /// <param name="value">The data object to use for populating the template. This can be an enumerable collection, DataTable, or other supported data source.</param>
    /// <param name="overwriteFile">If <c>true</c>, overwrites the file at the specified path, otherwise a <see cref="IOException"/> will be raised if the file already exists.</param>
    /// <param name="configuration">Optional configuration settings for the template fill operation.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    [CreateSyncVersion]
    public async Task FillTemplateAsync(string path, string templatePath, object value, bool overwriteFile = false,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var stream = overwriteFile ? File.Create(path) : File.Open(path, FileMode.CreateNew);
        await using var disposableStream = stream.ConfigureAwait(false); 

        await FillTemplateAsync(stream, templatePath, value, configuration, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Fills a template with the provided data and saves the result to a file.
    /// </summary>
    /// <param name="path">The destination file to write the filled document to.</param>
    /// <param name="templateStream">The stream containing the OpenXml template document.</param>
    /// <param name="value">The data object to use for populating the template. This can be an enumerable collection, DataTable, or other supported data source.</param>
    /// <param name="overwriteFile">If <c>true</c>, overwrites the file at the specified path, otherwise a <see cref="IOException"/> will be raised if the file already exists.</param>
    /// <param name="configuration">Optional configuration settings for the template fill operation.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    [CreateSyncVersion]
    public async Task FillTemplateAsync(string path, Stream templateStream, object value, bool  overwriteFile = false,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var stream = overwriteFile ? File.Create(path) : File.Open(path, FileMode.CreateNew);
        await using var disposableStream = stream.ConfigureAwait(false); 

        var template = GetOpenXmlTemplate(stream, configuration);
        await template.SaveAsByTemplateAsync(templateStream, value, cancellationToken).ConfigureAwait(false);
    }
    
    /// <summary>
    /// Fills a template with the provided data and saves the result to a destination stream.
    /// </summary>
    /// <param name="stream">The destination stream to write the filled document to.</param>
    /// <param name="templatePath">The path to the OpenXml template document.</param>
    /// <param name="value">The data object to use for populating the template. This can be an enumerable collection, DataTable, or other supported data source.</param>
    /// <param name="configuration">Optional configuration settings for the template fill operation.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    [CreateSyncVersion]
    public async Task FillTemplateAsync(Stream stream, string templatePath, object value,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var template = GetOpenXmlTemplate(stream, configuration);
        await template.SaveAsByTemplateAsync(templatePath, value, cancellationToken).ConfigureAwait(false);
    }
    
    /// <summary>
    /// Fills a template with the provided data and saves the result to a destination stream.
    /// </summary>
    /// <param name="stream">The destination stream to write the filled document to.</param>
    /// <param name="templateStream">The stream containing the OpenXml template document.</param>
    /// <param name="value">The data object to use for populating the template. This can be an enumerable collection, DataTable, or other supported data source.</param>
    /// <param name="configuration">Optional configuration settings for the template fill operation.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    [CreateSyncVersion]
    public async Task FillTemplateAsync(Stream stream, Stream templateStream, object value,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var template = GetOpenXmlTemplate(stream, configuration);
        await template.SaveAsByTemplateAsync(templateStream, value, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Fills a template with the provided data and saves the result to a file.
    /// </summary>
    /// <param name="path">The destination stream to write the filled document to.</param>
    /// <param name="templateBytes">A byte array containing the OpenXml template document.</param>
    /// <param name="value">The data object to use for populating the template. This can be an enumerable collection, DataTable, or other supported data source.</param>
    /// <param name="overwriteFile">If <c>true</c>, overwrites the file at the specified path, otherwise a <see cref="IOException"/> will be raised if the file already exists.</param>
    /// <param name="configuration">Optional configuration settings for the template fill operation.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    [CreateSyncVersion]
    public async Task FillTemplateAsync(string path, byte[] templateBytes, object value, bool overwriteFile = false,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var stream = overwriteFile ? File.Create(path) :  File.Open(path, FileMode.CreateNew);
        await using var disposableStream = stream.ConfigureAwait(false); 

        await FillTemplateAsync(stream, templateBytes, value, configuration, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Fills a template with the provided data and saves the result to a destination stream.
    /// </summary>
    /// <param name="stream">The destination stream to write the filled document to.</param>
    /// <param name="templateBytes">A byte array containing the OpenXml template document.</param>
    /// <param name="value">The data object to use for populating the template. This can be an enumerable collection, DataTable, or other supported data source.</param>
    /// <param name="configuration">Optional configuration settings for the template fill operation.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    [CreateSyncVersion]
    public async Task FillTemplateAsync(Stream stream, byte[] templateBytes, object value,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var template = GetOpenXmlTemplate(stream, configuration);
        await template.SaveAsByTemplateAsync(templateBytes, value, cancellationToken).ConfigureAwait(false);
    }

    #region Merge Cells

    /// <summary>
    /// Merges cells with identical values in a specified OpenXml document.
    /// </summary>
    /// <param name="mergedFilePath">The destination file to write the merged document to.</param>
    /// <param name="path">The file path to the original OpenXml document to process for cell merging.</param>
    /// <param name="configuration">Optional configuration settings for the merge operation.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    [CreateSyncVersion]
    public async Task MergeSameCellsAsync(string mergedFilePath, string path,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var stream = File.Create(mergedFilePath);
        await using var disposableStream = stream.ConfigureAwait(false); 

        await MergeSameCellsAsync(stream, path, configuration, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Merges cells with identical values in a specified OpenXml document.
    /// </summary>
    /// <param name="stream">The destination stream to write the merged document to.</param>
    /// <param name="path">The file path to the original OpenXml document to process for cell merging.</param>
    /// <param name="configuration">Optional configuration settings for the merge operation.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    [CreateSyncVersion]
    public async Task MergeSameCellsAsync(Stream stream, string path,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var template = GetOpenXmlTemplate(stream, configuration);
        await template.MergeSameCellsAsync(path, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Merges cells with identical values in a specified OpenXml document.
    /// </summary>
    /// <param name="stream">The destination stream to write the merged document to.</param>
    /// <param name="documentData">A byte array containing the original OpenXml document to process for cell merging.</param>
    /// <param name="configuration">Optional configuration settings for the merge operation.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    [CreateSyncVersion]
    public async Task MergeSameCellsAsync(Stream stream, byte[] documentData,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var template = GetOpenXmlTemplate(stream, configuration);
        await template.MergeSameCellsAsync(documentData, cancellationToken).ConfigureAwait(false);
    }

    
    private static OpenXmlTemplate GetOpenXmlTemplate(Stream stream, OpenXmlConfiguration? configuration)
    {
        var valueExtractor = new OpenXmlValueExtractor();
        return new OpenXmlTemplate(stream, configuration, valueExtractor);
    }

    #endregion
}
