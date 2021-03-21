<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <NuGetReference>System.Data.SQLite.Core</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>System.Data.SQLite</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Diagnostics</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Text</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	using (var connection = new SQLiteConnection("Data Source=:memory:"))
	{
		{
			var v = connection.Query(@"select 1 A,2 B ");
			var cV = v as IEnumerable<IDictionary<string,object>>;
			Console.WriteLine(v.GetType().FullName); // result : System.Collections.Generic.List`1[[Dapper.SqlMapper+DapperRow, Dapper, Version=2.0.0.0, Culture=neutral, PublicKeyToken=null]]
			Console.WriteLine(cV); // result : A=1,B=2
		}
		
		{
			var v = connection.Query(@"select 1 A,2 B ").ToList();
			var cV = v as IEnumerable<IDictionary<string, object>>;
			Console.WriteLine(v.GetType().FullName); // result : System.Collections.Generic.List`1[[System.Object, System.Private.CoreLib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=7cec85d7bea7798e]]
			Console.WriteLine(cV); // result : null
		}
		
		{
			var v = connection.Query(@"select 1 A,2 B ").ToList();
			foreach (IDictionary<string, object> e in v)
			{
				Console.WriteLine(v.GetType().FullName); // result : System.Collections.Generic.List`1[[System.Object, System.Private.CoreLib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=7cec85d7bea7798e]]
				Console.WriteLine(e); // result : A=1,B=2			
			}
		}
	}
	

}
