<Query Kind="Program">
  <Connection>
    <ID>2a2e3c0b-0e23-4992-bf19-66db2739e377</ID>
    <Persist>true</Persist>
    <Server>(localdb)\mssqllocaldb</Server>
    <Database>tempdb</Database>
  </Connection>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>Dapper</Namespace>
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
	{
		var list = new List<Dictionary<string, object>>() { new Dictionary<string, object>() };
		object obj = list;
		Console.WriteLine(obj is ICollection); //True
		Console.WriteLine(list.ToList() is ICollection); //True

		Console.WriteLine(obj is ICollection<IDictionary>); //False
		Console.WriteLine(list.ToList() is ICollection<IDictionary>); //False

		var type = list.GetType();
		var gt = type.GetGenericArguments()[0];
		Console.WriteLine(typeof(IDictionary).IsAssignableFrom(gt)); //True

	}

	using (var cn = Connection)
	{
		{
			object rows = cn.Query("select 1 id union all select 2");
			var type = rows.GetType();
			var gt = type.GetGenericArguments()[0]; //dapper row
			if(rows is IEnumerable && typeof(IDictionary<string,object>).IsAssignableFrom(gt))
			{
				foreach (IDictionary<string,object> row in rows as IEnumerable)
					Console.WriteLine(row);
			}
		}
		{
			var rows = cn.Query("select 1 id union all select 2").ToList();
			var type = rows.GetType();
			var gt = type.GetGenericArguments()[0]; //!!   object
		}
		{
			var rows = cn.Query("select 1 id union all select 2").ToList();
			var type = rows.GetType();
			var gt = type.GetGenericArguments()[0];

			//if (rows is ICollection && typeof(IDictionary<string,object>).IsAssignableFrom(gt))
			if (rows is IEnumerable && typeof(object) == (gt))
			{
				var values = (IEnumerable)rows;
				foreach (var row in values)
				{
					if (row != null)
					{
						gt = row.GetType();
						break;
					}
				}

				if (typeof(IDictionary<string, object>).IsAssignableFrom(gt)) //Dapper Rows
				{
					ICollection<string> keys = null;
					foreach (IDictionary<string, object> element in values)
					{
						keys = element.Keys;
						foreach (var key in keys)
						{
							var value = element[key];
							Console.WriteLine(value);
						}
					}
				}
			}
		}
	}
}

// You can define other methods, fields, classes and namespaces here
