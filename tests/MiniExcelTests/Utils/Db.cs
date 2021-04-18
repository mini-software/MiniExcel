/**
 This Class Modified from ExcelDataReader : https://github.com/ExcelDataReader/ExcelDataReader
 **/
namespace MiniExcelLibs.Tests.Utils
{
    using System.Data.SQLite;

    internal static class Db
    {
	   internal static SQLiteConnection GetConnection(string connectionString)
	   {
		  return new SQLiteConnection(connectionString);
	   }
    }

}
