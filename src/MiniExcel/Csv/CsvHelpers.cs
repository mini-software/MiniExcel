namespace MiniExcelLibs.Csv
{
    internal static class CsvHelpers
    {
        /// <summary>If content contains `;, "` then use "{value}" format</summary>
        public static string ConvertToCsvValue(string value)
        {
            if (value == null)
                return string.Empty;
            if (value.Contains("\""))
            {
                value = value.Replace("\"", "\"\"");
                return $"\"{value}\"";
            }
            else if (value.Contains(",") || value.Contains(" "))
            {
                return $"\"{value}\"";
            }
            return value;
        }
    }
}
