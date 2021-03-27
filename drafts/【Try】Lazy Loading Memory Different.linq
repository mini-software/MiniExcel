<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>Xunit</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
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

[Fact]
void AsICollectionTest()
{
	// lazy loading can avoid loading all data into memory first 
	Console.WriteLine("==== AsICollectionTest ====");
	Console.WriteLine($"start memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");
	var values = Enumerable.Range(1, 10000000).Select((s, index) => new { index, value = Guid.NewGuid() }) as IEnumerable;
	var index = 0;
	foreach (var element in values)
	{
		index++;
		if (index % 1000000 == 0 || index == 1)
			Console.WriteLine($"no.{index} memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");
	}
	/*
	start memory usage: 54 MB
	no.1 memory usage: 54 MB
	no.1000000 memory usage: 54 MB
	no.2000000 memory usage: 54 MB
	no.3000000 memory usage: 54 MB
	no.4000000 memory usage: 54 MB
	no.5000000 memory usage: 54 MB
	no.6000000 memory usage: 54 MB
	no.7000000 memory usage: 54 MB
	no.8000000 memory usage: 54 MB
	no.9000000 memory usage: 54 MB
	no.10000000 memory usage: 54 MB
	*/
}

void Main2()
{
	// lazy loading can avoid loading all data into memory first 
	Console.WriteLine($"start memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");
	var values = Enumerable.Range(1, 10000000).Select((s, index) => new { index, value = Guid.NewGuid() });
	var index = 0;
	foreach (var element in values)
	{
		index++;
		if (index % 1000000 == 0 || index == 1)
			Console.WriteLine($"no.{index} memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");
	}
	/*
start memory usage: 54 MB
no.1 memory usage: 54 MB
no.1000000 memory usage: 54 MB
no.2000000 memory usage: 54 MB
no.3000000 memory usage: 54 MB
no.4000000 memory usage: 54 MB
no.5000000 memory usage: 54 MB
no.6000000 memory usage: 54 MB
no.7000000 memory usage: 54 MB
no.8000000 memory usage: 54 MB
no.9000000 memory usage: 54 MB
no.10000000 memory usage: 54 MB
	*/
}

void Main()
{
	//RunTests();  // Call RunTests() or press Alt+Shift+T to initiate testing.

	Console.WriteLine($"start memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");
	var values = Enumerable.Range(1, 10000000).Select((s, index) => new { index, value = Guid.NewGuid() }).ToList();
	Console.WriteLine($"end memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");
	/*
start memory usage: 54 MB
end memory usage: 572 MB
	*/
}

// You can define other methods, fields, classes and namespaces here
