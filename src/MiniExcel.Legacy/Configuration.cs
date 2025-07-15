using MiniExcelLib;

namespace MiniExcelLibs;


public interface IConfiguration : IMiniExcelConfiguration;

public sealed class OpenXmlConfiguration : MiniExcelLib.OpenXml.OpenXmlConfiguration, IConfiguration;
public sealed class CsvConfiguration : MiniExcelLib.Csv.CsvConfiguration, IConfiguration;