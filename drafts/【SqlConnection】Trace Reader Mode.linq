<Query Kind="Program">
  <Connection>
    <ID>2a2e3c0b-0e23-4992-bf19-66db2739e377</ID>
    <Persist>true</Persist>
    <Server>(localdb)\mssqllocaldb</Server>
    <Database>Northwind</Database>
  </Connection>
  <NuGetReference>AngleSharp</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

void Main()
{
	using (var cn = (SqlConnection)this.Connection)
	{
		cn.Open();
		using(var cmd = cn.CreateCommand()){
			cmd.CommandText="select 1 id";
			using(var reader = cmd.ExecuteReader()){
				var name = reader.GetName(0);
				Console.WriteLine(name);
				var dt = reader.GetSchemaTable();
				Console.WriteLine(dt);
				var v = reader.GetValue(0);
				var vs = reader.GetValues(null);
			}
		}
	}
}

// You can define other methods, fields, classes and namespaces here
