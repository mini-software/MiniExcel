namespace MiniExcelLibs.Tests.Utils
{
    using System.Collections.Generic;
    using System;
    using System.Data.SQLite;
    using System.Text;

    internal static class Db
    {
        internal static SQLiteConnection GetConnection(string connectionString = "Data Source=:memory:")
        {
            return new SQLiteConnection(connectionString);
        }

        internal static string GenerateDummyQuery(List<Dictionary<string, object>> data)
        {
            if (data == null || data.Count == 0)
                throw new ArgumentException("The data list cannot be null or empty.");

            var queryBuilder = new StringBuilder();

            for (int i = 0; i < data.Count; i++)
            {
                var row = data[i];
                var selectStatement = new StringBuilder("SELECT ");

                foreach (var kvp in row)
                {
                    string columnName = kvp.Key;
                    object value = kvp.Value;

                    // Format value based on its type
                    string formattedValue = value switch
                    {
                        string str => $"'{str.Replace("'", "''")}'", // Escape single quotes in strings
                        DateTime dt => $"'{dt:yyyy-MM-dd HH:mm:ss}'", // Format datetime as string
                        bool b => b ? "1" : "0", // Convert boolean to 1 or 0
                        _ => value.ToString() // Use value as-is for numbers and other types
                    };

                    selectStatement.Append($"{formattedValue} AS {columnName}, ");
                }

                // Remove the trailing comma and space
                selectStatement.Length -= 2;

                // Add UNION ALL between each row, except for the last one
                if (i < data.Count - 1)
                    selectStatement.Append(" UNION ALL ");

                queryBuilder.AppendLine(selectStatement.ToString());
            }

            return queryBuilder.ToString();
        }
    }
}
