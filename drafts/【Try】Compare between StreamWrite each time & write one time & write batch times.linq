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
	Test1();
	GC.Collect();
	Test2();
	GC.Collect();
	Test3();
	GC.Collect();
	/*
==== write each time ====
memory usage: 61 MB , 3ms
memory usage: 59 MB , 2908ms
memory usage: 59 MB , 2916ms
==== write one time ====
memory usage: 59 MB , 3ms
memory usage: 785 MB , 2874ms
memory usage: 1472 MB , 3817ms
==== write batch ====
memory usage: 62 MB , 3ms
memory usage: 62 MB , 8ms
memory usage: 129 MB , 336ms
memory usage: 265 MB , 725ms
memory usage: 402 MB , 1121ms
memory usage: 544 MB , 1511ms
memory usage: 686 MB , 1876ms
memory usage: 829 MB , 2243ms
memory usage: 973 MB , 2608ms
memory usage: 195 MB , 3104ms
memory usage: 226 MB , 3523ms
memory usage: 196 MB , 3941ms
memory usage: 264 MB , 4051ms
	*/
}

[Fact]
void Test1()
{
	Console.WriteLine("==== write each time ====");
	var sw = new Stopwatch();
	sw.Start();
	//RunTests();  // Call RunTests() or press Alt+Shift+T to initiate testing.
	Console.WriteLine($"memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB , {sw.ElapsedMilliseconds}ms");
	// compare write each time or write one time or write batch
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.txt");
	using (var stream = File.Create(path))
	using (var writer = new StreamWriter(stream))
	{
		for (int i = 0; i < 10000000; i++)
			writer.Write(Guid.NewGuid().ToString());
		Console.WriteLine($"memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB , {sw.ElapsedMilliseconds}ms");
	}
	Console.WriteLine($"memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB , {sw.ElapsedMilliseconds}ms");
	File.Delete(path);
}

[Fact]
void Test2()
{
	Console.WriteLine("==== write one time ====");
	var sw = new Stopwatch();
	sw.Start();
	//RunTests();  // Call RunTests() or press Alt+Shift+T to initiate testing.
	Console.WriteLine($"memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB , {sw.ElapsedMilliseconds}ms");
	// compare write each time or write one time or write batch
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.txt");
	using (var stream = File.Create(path))
	using (var writer = new StreamWriter(stream))
	{
		var sb = new StringBuilder();
		for (int i = 0; i < 10000000; i++)
			sb.Append(Guid.NewGuid().ToString());
		Console.WriteLine($"memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB , {sw.ElapsedMilliseconds}ms");
		writer.Write(sb.ToString());
	}
	Console.WriteLine($"memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB , {sw.ElapsedMilliseconds}ms");
	File.Delete(path);
}

[Fact]
void Test3()
{
	Console.WriteLine("==== write batch ====");
	var sw = new Stopwatch();
	sw.Start();
	//RunTests();  // Call RunTests() or press Alt+Shift+T to initiate testing.
	Console.WriteLine($"memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB , {sw.ElapsedMilliseconds}ms");
	// compare write each time or write one time or write batch
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.txt");
	using (var stream = File.Create(path))
	using (var writer = new StreamWriter(stream))
	{
		var sb = new StringBuilder();
		for (int i = 0; i < 10000000; i++)
		{
			sb.Append(Guid.NewGuid().ToString());
			if (i % 1000000 == 0)
			{
				Console.WriteLine($"memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB , {sw.ElapsedMilliseconds}ms");
				writer.Write(sb.ToString());
				sb = null;
				sb = new StringBuilder();
			}
		}
		Console.WriteLine($"memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB , {sw.ElapsedMilliseconds}ms");
		writer.Write(sb.ToString());
	}
	Console.WriteLine($"memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB , {sw.ElapsedMilliseconds}ms");
	File.Delete(path);
}

// You can define other methods, fields, classes and namespaces here
