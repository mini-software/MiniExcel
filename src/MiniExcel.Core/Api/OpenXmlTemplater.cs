using MiniExcelLib.Core.OpenXml.Picture;
using MiniExcelLib.Core.OpenXml.Templates;

// ReSharper disable once CheckNamespace
namespace MiniExcelLib.Core;

public sealed partial class OpenXmlTemplater
{
    internal OpenXmlTemplater() { }
    
    [CreateSyncVersion]
    public async Task AddPictureAsync(string path, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
    {
        using var stream = File.Open(path, FileMode.OpenOrCreate);
        await MiniExcelPictureImplement.AddPictureAsync(stream, cancellationToken, images).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task AddPictureAsync(Stream excelStream, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
    {
        await MiniExcelPictureImplement.AddPictureAsync(excelStream, cancellationToken, images).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task ApplyTemplateAsync(string path, string templatePath, object value,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        using var stream = File.Create(path);
        await ApplyTemplateAsync(stream, templatePath, value, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task ApplyTemplateAsync(string path, Stream templateStream, object value,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        using var stream = File.Create(path);
        var template = GetOpenXmlTemplate(stream, configuration);
        await template.SaveAsByTemplateAsync(templateStream, value, cancellationToken).ConfigureAwait(false);
    }
    
    [CreateSyncVersion]
    public async Task ApplyTemplateAsync(Stream stream, string templatePath, object value,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var template = GetOpenXmlTemplate(stream, configuration);
        await template.SaveAsByTemplateAsync(templatePath, value, cancellationToken).ConfigureAwait(false);
    }
    
    [CreateSyncVersion]
    public async Task ApplyTemplateAsync(Stream stream, Stream templateStream, object value,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var template = GetOpenXmlTemplate(stream, configuration);
        await template.SaveAsByTemplateAsync(templateStream, value, cancellationToken).ConfigureAwait(false);
    }
    
    [CreateSyncVersion]
    public async Task ApplyTemplateAsync(string path, byte[] templateBytes, object value,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        using var stream = File.Create(path);
        await ApplyTemplateAsync(stream, templateBytes, value, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task ApplyTemplateAsync(Stream stream, byte[] templateBytes, object value,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var template = GetOpenXmlTemplate(stream, configuration);
        await template.SaveAsByTemplateAsync(templateBytes, value, cancellationToken).ConfigureAwait(false);
    }

    #region Merge Cells

    [CreateSyncVersion]
    public async Task MergeSameCellsAsync(string mergedFilePath, string path,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        using var stream = File.Create(mergedFilePath);
        await MergeSameCellsAsync(stream, path, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task MergeSameCellsAsync(Stream stream, string path,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var template = GetOpenXmlTemplate(stream, configuration);
        await template.MergeSameCellsAsync(path, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task MergeSameCellsAsync(Stream stream, byte[] file,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var template = GetOpenXmlTemplate(stream, configuration);
        await template.MergeSameCellsAsync(file, cancellationToken).ConfigureAwait(false);
    }

    
    private static OpenXmlTemplate GetOpenXmlTemplate(Stream stream, OpenXmlConfiguration? configuration)
    {
        var valueExtractor = new OpenXmlValueExtractor();
        return new OpenXmlTemplate(stream, configuration, valueExtractor);
    }

    #endregion
}