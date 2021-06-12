namespace MiniExcelLibs.Tests.Utils
{
    using System.Data.SQLite;

    internal static class Db
    {
	   internal static SQLiteConnection GetConnection(string connectionString= "Data Source=:memory:")
	   {
		  return new SQLiteConnection(connectionString);
	   }
    }

}
