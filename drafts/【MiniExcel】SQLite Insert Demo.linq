<Query Kind="Program">
  <NuGetReference>AngleSharp</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>Dapper.Contrib</NuGetReference>
  <NuGetReference>Imgur.API</NuGetReference>
  <NuGetReference>LinqToExcel</NuGetReference>
  <NuGetReference>LinqToExcel_x64</NuGetReference>
  <NuGetReference>Markdig</NuGetReference>
  <NuGetReference Version="0.0.6-beta" Prerelease="true">MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>NPOI</NuGetReference>
  <NuGetReference>NPOI.Extension</NuGetReference>
  <NuGetReference>OfficeOpenXml.Core.ExcelPackage</NuGetReference>
  <NuGetReference>Oracle.ManagedDataAccess</NuGetReference>
  <NuGetReference>Quartz</NuGetReference>
  <NuGetReference>RestSharp</NuGetReference>
  <NuGetReference>SqlKata</NuGetReference>
  <NuGetReference>System.Data.SQLite.Core</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SQLite</Namespace>
  <AppConfig OverrideConnection="true">
    <Content>
      <configuration>
        <connectionStrings>
          <add name="Oracle" connectionString="Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.2.16)(PORT = 1521)) (CONNECT_DATA = (SERVICE_NAME=MES)) );USER ID=MES;PASSWORD=MES;Min Pool Size =1;Max Pool Size = 50;" providerName="System.Data.OracleClient" />
          <add name="ERP" connectionString="Data Source=192.168.1.2;    Initial Catalog=DB_KENT;Persist Security Info=True;    User ID=sa;Password=123456" />
          <add name="SQLiteConn" connectionString="Data Source=E:\git\ITWeiHanDEV\LINQPad Queries\ITWeiHanWeb_Projects\test.db;FailIfMissing=True" />
        </connectionStrings>
      </configuration>
    </Content>
  </AppConfig>
</Query>

void Main()
{
	var path = @"D:\git\MiniExcel\samples\xlsx\Test1,000,000x10\Test1,000,000x10.xlsx";
	var tempSqlitePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
	var connectionString = $"Data Source={tempSqlitePath};Version=3;";
	SQLiteConnection.CreateFile(tempSqlitePath);

	using (var connection = new SQLiteConnection(connectionString))
	{
		connection.Execute(@"create table T (A varchar(20),B varchar(20));");
	}

	Stopwatch sw = new Stopwatch();
	sw.Start();
	Console.WriteLine("start memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB");

	using (var connection = new SQLiteConnection(connectionString))
	{
		connection.Open();
		using (var transaction = connection.BeginTransaction())
		using (var stream = File.OpenRead(path))
		{
			var rows = stream.Query();
			foreach (var row in rows)
				connection.Execute("insert into T (A,B) values (@A,@B)", new { row.A, row.B }, transaction: transaction);
			transaction.Commit();
		}
		
		Console.WriteLine("end memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");
	}

	using (var connection = new SQLiteConnection(connectionString))
	{
		var count = connection.ExecuteScalar<int>("select count(*) from T");
		Console.WriteLine($"count : {count}");
	}
}

