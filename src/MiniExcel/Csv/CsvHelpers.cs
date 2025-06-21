namespace MiniExcelLibs.Csv;

internal static class CsvHelpers
{
    /// <summary>If content contains special characters then use "{value}" format</summary>
    public static string ConvertToCsvValue(string? value, CsvConfiguration configuration)
    {
        if (value is null)
            return string.Empty;

        if (value.Contains("\""))
        {
            value = value.Replace("\"", "\"\"");
            return $"\"{value}\"";
        }
            
        var shouldQuote = configuration.AlwaysQuote ||
                          (configuration.QuoteWhitespaces && value.Contains(" ")) ||
                          value.Contains(configuration.Seperator.ToString()) ||
                          value.Contains("\r") ||
                          value.Contains("\n");
            
        return shouldQuote ? $"\"{value}\"" : value;
    }
}