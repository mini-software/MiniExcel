namespace MiniExcelLibs.Csv
{
    internal static class CsvHelpers
    {
        /// <summary>If content contains special characters then use "{value}" format</summary>
        public static string ConvertToCsvValue(string value, bool alwaysQuote, char separator)
        {
            if (value == null)
                return string.Empty;

            if (value.Contains("\""))
            {
                value = value.Replace("\"", "\"\"");
                return $"\"{value}\"";
            }

            if (value.Contains(separator.ToString()) || value.Contains(" ") || value.Contains("\n") || value.Contains("\r"))
            {
                return $"\"{value}\"";
            }

            if (alwaysQuote)
                return $"\"{value}\"";
 
            return value;
        }
    }
}
