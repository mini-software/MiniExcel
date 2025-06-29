using MiniExcelLib.Core.OpenXml;
using Zomp.SyncMethodGenerator;

namespace MiniExcelLib;

public static partial class MiniExcel
{
    public static partial class Templater
    {
        [CreateSyncVersion]
        public static async Task ApplyXlsxTemplateAsync(string path, string templatePath, object value,
            OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
        {
            using var stream = File.Create(path);
            await ApplyXlsxTemplateAsync(stream, templatePath, value, configuration, cancellationToken).ConfigureAwait(false);
        }

        [CreateSyncVersion]
        public static async Task ApplyXlsxTemplateAsync(string path, byte[] templateBytes, object value,
            OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
        {
            using var stream = File.Create(path);
            await ApplyXlsxTemplateAsync(stream, templateBytes, value, configuration, cancellationToken).ConfigureAwait(false);
        }

        [CreateSyncVersion]
        public static async Task ApplyXlsxTemplateAsync(Stream stream, string templatePath, object value,
            OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
        {
            var template = GetOpenXmlTemplate(stream, configuration);
            await template.SaveAsByTemplateAsync(templatePath, value, cancellationToken).ConfigureAwait(false);
        }

        [CreateSyncVersion]
        public static async Task ApplyXlsxTemplateAsync(Stream stream, byte[] templateBytes, object value,
            OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
        {
            var template = GetOpenXmlTemplate(stream, configuration);
            await template.SaveAsByTemplateAsync(templateBytes, value, cancellationToken).ConfigureAwait(false);
        }

        #region Merge Cells

        [CreateSyncVersion]
        public static async Task MergeSameCellsAsync(string mergedFilePath, string path,
            OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
        {
            using var stream = File.Create(mergedFilePath);
            await MergeSameCellsAsync(stream, path, configuration, cancellationToken).ConfigureAwait(false);
        }

        [CreateSyncVersion]
        public static async Task MergeSameCellsAsync(Stream stream, string path,
            OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
        {
            var template = GetOpenXmlTemplate(stream, configuration);
            await template.MergeSameCellsAsync(path, cancellationToken).ConfigureAwait(false);
        }

        [CreateSyncVersion]
        public static async Task MergeSameCellsAsync(Stream stream, byte[] file,
            OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
        {
            var template = GetOpenXmlTemplate(stream, configuration);
            await template.MergeSameCellsAsync(file, cancellationToken).ConfigureAwait(false);
        }

        #endregion
    }
}