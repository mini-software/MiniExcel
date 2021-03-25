<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <NuGetReference>System.Data.SQLite.Core</NuGetReference>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Diagnostics</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	Test.Program.Main();
}

// You can define other methods, fields, classes and namespaces here
namespace Test
{
	using System;
	using System.IO;
	using MiniExcelLibs;
	using System.Linq;
	using Dapper;
	using System.Data.SQLite;
	public class Program
	{
		public static void Main()
		{
			var path = "demo.xlsx";
			var values = new[] { new { A = "Github", B = DateTime.Parse("2021-01-01") }, new { A = "Microsoft", B = DateTime.Parse("2021-02-01") } };

			// create 
			using (var stream = File.Create(path))
				stream.SaveAs(values);

			// query dynamic
			using (var stream = File.OpenRead(path))
			{
				var rows = stream.Query(useHeaderRow: true).ToList();
				Console.WriteLine(rows[0].A); //Github
				Console.WriteLine(rows[0].B); //2021-01-01 12:00:00 AM
				Console.WriteLine(rows[1].A); //Microsoft
				Console.WriteLine(rows[1].B); //2021-02-01 12:00:00 AM		
			}

			// query type mapping
			using (var stream = File.OpenRead(path))
			{
				var rows = stream.Query<Demo>().ToList();
				Console.WriteLine(rows[0].A); //Github
				Console.WriteLine(rows[0].B); //2021-01-01 12:00:00 AM
				Console.WriteLine(rows[1].A); //Microsoft
				Console.WriteLine(rows[1].B); //2021-02-01 12:00:00 AM			
			}

			File.Delete(path);

			// Create by DapperRows
			using (var connection = new System.Data.SQLite.SQLiteConnection("Data Source=:memory:"))
			{
				var rows = connection.Query(@"select 'MiniExcel' as Column1,1 as Column2 union all select 'Github',2");
				MiniExcel.SaveAs(path, rows);
			}

			using (var stream = File.OpenRead(path))
			{
				var rows = stream.Query(useHeaderRow:true).ToList();
				Console.WriteLine(rows[0].Column1); //MiniExcel
				Console.WriteLine(rows[0].Column2); //1
				Console.WriteLine(rows[1].Column1); //Github
				Console.WriteLine(rows[1].Column2); //2			
			}
		}

		public class Demo
		{
			public string A { get; set; }
			public string B { get; set; }
		}
	}
}