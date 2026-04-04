using MiniExcelLib.OpenXml;
using MiniExcelLib.OpenXml.Picture;
using MiniExcelLib.OpenXml.Templates;
using OpenXmlTemplate = MiniExcelLib.OpenXml.Templates.OpenXmlTemplate;

// ReSharper disable once CheckNamespace
namespace MiniExcelLib.Core;

public sealed partial class OpenXmlTemplater
{
    internal OpenXmlTemplater() { }
    
    [CreateSyncVersion]
    public async Task AddPictureAsync(string path, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
    {
#if NET8_0_OR_GREATER
        var stream = File.Open(path, FileMode.OpenOrCreate);
        await using var disposableStream = stream.ConfigureAwait(false); 
#else
        using var stream = File.Open(path, FileMode.OpenOrCreate);
#endif
        await MiniExcelPictureImplement.AddPictureAsync(stream, cancellationToken, images).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task AddPictureAsync(Stream excelStream, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
    {
        await MiniExcelPictureImplement.AddPictureAsync(excelStream, cancellationToken, images).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task FillTemplateAsync(string path, string templatePath, object value, bool overwriteFile = false,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
#if NET8_0_OR_GREATER
        var stream = overwriteFile ? File.Create(path) : File.Open(path, FileMode.CreateNew);
        await using var disposableStream = stream.ConfigureAwait(false); 
#else
        using var stream = overwriteFile ? File.Create(path) : File.Open(path, FileMode.CreateNew);
#endif
        await FillTemplateAsync(stream, templatePath, value, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task FillTemplateAsync(string path, Stream templateStream, object value, bool  overwriteFile = false,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
#if NET8_0_OR_GREATER
        var stream = overwriteFile ? File.Create(path) : File.Open(path, FileMode.CreateNew);
        await using var disposableStream = stream.ConfigureAwait(false); 
#else
        using var stream = overwriteFile ? File.Create(path) : File.Open(path, FileMode.CreateNew);
#endif

        var template = GetOpenXmlTemplate(stream, configuration);
        await template.SaveAsByTemplateAsync(templateStream, value, cancellationToken).ConfigureAwait(false);
    }
    
    [CreateSyncVersion]
    public async Task FillTemplateAsync(Stream stream, string templatePath, object value,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var template = GetOpenXmlTemplate(stream, configuration);
        await template.SaveAsByTemplateAsync(templatePath, value, cancellationToken).ConfigureAwait(false);
    }
    
    [CreateSyncVersion]
    public async Task FillTemplateAsync(Stream stream, Stream templateStream, object value,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var template = GetOpenXmlTemplate(stream, configuration);
        await template.SaveAsByTemplateAsync(templateStream, value, cancellationToken).ConfigureAwait(false);
    }
    
    [CreateSyncVersion]
    public async Task FillTemplateAsync(string path, byte[] templateBytes, object value, bool overwriteFile = false,
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
#if NET8_0_OR_GREATER
        var stream = overwriteFile ? File.Create(path) :  File.Open(path, FileMode.CreateNew);
        await using var disposableStream = stream.ConfigureAwait(false); 
#else
        using var stream = overwriteFile ? File.Create(path) :  File.Open(path, FileMode.CreateNew);
#endif
        await FillTemplateAsync(stream, templateBytes, value, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task FillTemplateAsync(Stream stream, byte[] templateBytes, object value,
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
#if NET8_0_OR_GREATER
        var stream = File.Create(mergedFilePath);
        await using var disposableStream = stream.ConfigureAwait(false); 
#else
        using var stream = File.Create(mergedFilePath);
#endif
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