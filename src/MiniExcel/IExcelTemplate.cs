using System.Threading.Tasks;

namespace MiniExcelLibs
{
    internal interface IExcelTemplate
    {
        void SaveAsByTemplate(string templatePath, object value);
        void SaveAsByTemplate(byte[] templateBtyes, object value);
    }

    internal interface IExcelTemplateAsync : IExcelTemplate
    {
        Task SaveAsByTemplateAsync(string templatePath, object value);
        Task SaveAsByTemplateAsync(byte[] templateBtyes, object value);
    }
}
