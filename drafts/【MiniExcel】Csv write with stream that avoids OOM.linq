<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Xunit</Namespace>
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

#load "xunit"

[Fact()]
void StringBuildWriteAllText_Test()
{
	Console.WriteLine("==== StringBuildWriteAllText_Test ====");
	Console.WriteLine("start memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB");
	var input = Enumerable.Range(1, 1000000).Select(s => Guid.NewGuid().ToString()).ToList();
	Console.WriteLine("memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB");
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.csv");
	
	var sb = String.Join(Environment.NewLine,input);
	Console.WriteLine("memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB");
	File.WriteAllText(path, sb.ToString());
	Console.WriteLine("end memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB");
	
	Console.WriteLine(path);
}

[Fact()]
void StreamWrite_Test()
{
	Console.WriteLine("==== StreamWrite_Test ====");
	Console.WriteLine("start memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB");
	var input = Enumerable.Range(1,1000000).Select(s=> Guid.NewGuid().ToString()).ToList();
	Console.WriteLine("memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB");
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.csv");
	
	using (var stream = File.CreateText(path))
	{
		var i = 0;
		foreach (var e in input)
		{
			i++;
			if (i % 100000 == 0)
				Console.WriteLine($"{i}. memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB");
			stream.Write(e + Environment.NewLine);
		}
		Console.WriteLine("end memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB");
	}
	Console.WriteLine(path);
}

// You can define other methods, fields, classes and namespaces here

