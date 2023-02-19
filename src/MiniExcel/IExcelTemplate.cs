using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    internal interface IExcelTemplate
    {
        void SaveAsByTemplate(string templatePath, object value);
        void SaveAsByTemplate(byte[] templateBtyes, object value);
        void MergeSameCells(string path);
    }

    internal interface IExcelTemplateAsync : IExcelTemplate
    {
        Task SaveAsByTemplateAsync(string templatePath, object value,CancellationToken cancellationToken = default(CancellationToken));
        Task SaveAsByTemplateAsync(byte[] templateBtyes, object value,CancellationToken cancellationToken = default(CancellationToken));
        Task MergeSameCellsAsync(string path, CancellationToken cancellationToken = default(CancellationToken));
    }
}
