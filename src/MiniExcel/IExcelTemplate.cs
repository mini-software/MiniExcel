namespace MiniExcelLibs
{
    internal interface IExcelTemplate
    {
        void SaveAsByTemplate(string templatePath, object value);
        void SaveAsByTemplate(byte[] templateBtyes, object value);
    }
}
