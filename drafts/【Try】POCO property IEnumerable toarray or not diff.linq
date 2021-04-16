<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
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

void Main()
{
	var st = new Stopwatch();
	st.Start();
	Console.WriteLine($"time : {st.ElapsedMilliseconds} ms");
	Console.WriteLine($"memory usage : {Process.GetCurrentProcess().WorkingSet64/(1024*1024)} MB");
	
	var MyClass = new MyClass();
	MyClass.CellIEnumerableValues = Enumerable.Range(1,100_000_000).Select(s=>"psjdfpsdfjpsdjfpj2p3jp4j23pj4p23j4pj32p4j23p4fspdjfpsdjpfjspdfjpsdfjpsodjfpsodjfposjdpfojsdpfojspdofjspdofjpsdojfposdjfpsdjf");
	
	Console.WriteLine($"time : {st.ElapsedMilliseconds} ms");
	Console.WriteLine($"memory usage : {Process.GetCurrentProcess().WorkingSet64/(1024*1024)} MB");
	
	for (int i = 0; i < 3; i++)
	{
		var idx = 0;
		foreach (var element in MyClass.CellIEnumerableValues)
		{
			//Console.WriteLine(element);
			idx++;
		}
		Console.WriteLine($"count : {idx}");
		Console.WriteLine($"time : {st.ElapsedMilliseconds} ms");
		Console.WriteLine($"memory usage : {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");
	}
}

/*
Non-ToArray result:
	time : 0 ms
	memory usage : 52 MB
	time : 5 ms
	memory usage : 52 MB
	count : 100000000
	time : 799 ms
	memory usage : 52 MB
	count : 100000000
	time : 1588 ms
	memory usage : 52 MB
	count : 100000000
	time : 2369 ms
	memory usage : 52 MB

ToArray Result:
	time : 0 ms
	memory usage : 49 MB
	time : 751 ms
	memory usage : 813 MB
	count : 100000000
	time : 2305 ms
	memory usage : 813 MB
	count : 100000000
	time : 3819 ms
	memory usage : 813 MB
	count : 100000000
	time : 5318 ms
	memory usage : 813 MB


summary:
System will not save IEnumerable to memory if type = IEnumerable
*/

public decimal GetCurrentMemoryUsage() => Process.GetCurrentProcess().WorkingSet64/(1024*1024);

// You can define other methods, fields, classes and namespaces here
class MyClass
{
	public IEnumerable CellIEnumerableValues { get; set; }
}