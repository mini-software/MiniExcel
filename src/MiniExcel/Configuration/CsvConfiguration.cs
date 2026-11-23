// ReSharper disable once CheckNamespace
namespace MiniExcelLibs.Csv;

public sealed class CsvConfiguration : MiniExcelLib.Csv.CsvConfiguration, IConfiguration
{
    public bool ReadEmptyStringAsNull
    {
        get => ReadEmptyFieldsAsDefault;
        set => ReadEmptyFieldsAsDefault = value;
    }
}
