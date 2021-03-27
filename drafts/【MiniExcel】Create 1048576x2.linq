<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Xunit</Namespace>
  <RemoveNamespace>System.Collections</RemoveNamespace>
  <RemoveNamespace>System.Collections.Generic</RemoveNamespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Reflection</RemoveNamespace>
  <RemoveNamespace>System.Text</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

#load "xunit"

void Main()
{
	Stopwatch sw = new Stopwatch();
	sw.Start();
	Console.WriteLine($"start memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");

	var values = Enumerable.Range(1, 1048575).Select((s,index) => new {value1=Guid.NewGuid(),value2=Guid.NewGuid()});
	Console.WriteLine($"memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");

	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
	using (var stream = File.Create(path))
	{
		stream.SaveAs(values);
	}
}

