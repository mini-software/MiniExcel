<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
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

namespace Test
{
	using System;
	using System.IO;
	using System.Text;
	using Newtonsoft.Json;
	using MiniExcelLibs;
	public class Program
	{
		public static void Main()
		{
			// SaveAs & Create by MemoryStream
			Byte[] bytes = null;
			using (var stream = new MemoryStream())
			{
				var value = new[] { new { A = "A1", B = "B1" }, new { A = "A2", B = "B2" } };
				stream.SaveAs(value: value, excelType: ExcelType.CSV);

				bytes = stream.ToArray();
				Console.WriteLine(Encoding.UTF8.GetString(bytes));
				/*result:
					A,B
					A1,B1
					A2,B2
				*/
			}

			using (var stream = new MemoryStream(bytes))
			{
				var rows = stream.Query(useHeaderRow: true);
				Console.WriteLine(JsonConvert.SerializeObject(rows)); // result : [{"A":"A1","B":"B1"},{"A":"A2","B":"B2"}]
			}
		}
	}
}
// You can define other methods, fields, classes and namespaces here
