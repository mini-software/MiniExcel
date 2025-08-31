using MiniExcelLib.Core;


namespace MiniExcelLibs
{
    public interface IConfiguration : IMiniExcelConfiguration;
}

namespace MiniExcelLibs.OpenXml
{
    public sealed class OpenXmlConfiguration : MiniExcelLib.Core.OpenXml.OpenXmlConfiguration, IConfiguration;
}

namespace MiniExcelLibs.Csv
{
    public sealed class CsvConfiguration : MiniExcelLib.Csv.CsvConfiguration, IConfiguration;
}
