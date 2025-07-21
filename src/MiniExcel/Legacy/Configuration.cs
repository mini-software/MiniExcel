using MiniExcelLib.Core;

namespace MiniExcelLib.Legacy;


public interface IConfiguration : IMiniExcelConfiguration;

public sealed class OpenXmlConfiguration : Core.OpenXml.OpenXmlConfiguration, IConfiguration;
public sealed class CsvConfiguration : MiniExcelLib.Csv.CsvConfiguration, IConfiguration;