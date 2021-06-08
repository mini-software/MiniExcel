/**
 This Class Modified from ExcelDataReader : https://github.com/ExcelDataReader/ExcelDataReader
 **/
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
