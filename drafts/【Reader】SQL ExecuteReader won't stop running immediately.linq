<Query Kind="Program">
  <Connection>
    <ID>2a2e3c0b-0e23-4992-bf19-66db2739e377</ID>
    <Persist>true</Persist>
    <Server>(localdb)\mssqllocaldb</Server>
    <Database>Northwind</Database>
  </Connection>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>Dapper</Namespace>
  <Namespace>System.Dynamic</Namespace>
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
	Test1();
}


void TestDapper()
{
	using (var connection = this.Connection)
	{
		foreach (var item in connection.Query("select t1.* from  master..spt_values t1, master..spt_values t2"))
		{
			// dapper default not support immdiately load
		}
	}
}

void Test1()
{
	foreach (var element in GetReaderValues())
	{
		Console.WriteLine(element);
	}
}

IEnumerable<dynamic> GetReaderValues()
{
	using (var connection = this.Connection)
	{
		connection.Open();
		var cmd = connection.CreateCommand();
		cmd.CommandText = "select t1.* from  master..spt_values t1, master..spt_values t2";

		using (var reader = cmd.ExecuteReader(behavior: CommandBehavior.SequentialAccess | CommandBehavior.SingleResult))
		{
			while (reader.Read())
			{
				var d = new ExpandoObject();
				for (int i = 0; i < reader.FieldCount; i++)
				{
					d.TryAdd(reader.GetName(i), reader.GetValue(i));
				}
				yield return d;
			}
		} //it won't stop running immediately
	}
}

// CommandBehavior.SchemaOnly is most quicky
void Test2()
{
	using (var connection = this.Connection)
	{
		connection.Open();
		var cmd = connection.CreateCommand();
		cmd.CommandText = "select t1.* from  master..spt_values t1, master..spt_values t2";
		using (var reader = cmd.ExecuteReader(behavior: CommandBehavior.SchemaOnly))
		{
			var dt = reader.GetSchemaTable();
			Console.WriteLine(dt);

			while (reader.Read()) { } // it will not close query immediately
			connection.Close();
			reader.DisposeAsync().GetAwaiter().GetResult();
		}
	}
}