<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.IO.Compression</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

//result : ![](https://i.imgur.com/QtQAteb.png)
void Main2() 
{
	Stopwatch sw = new Stopwatch();
	sw.Start();
	Console.WriteLine("start memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB");

	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.zip");
	using (var stream = File.Create(path))
	using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true, Utf8WithBom))
	{
		var sb = new StringBuilder();
		for (var i = 0; i <= 10000000; i++)
		{
			sb.AppendLine("Hello World\n\r");
			if (i % 1000000 == 0)
				Console.WriteLine($"time.{i} memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");
		}

		var entry = archive.CreateEntry("test.txt");
		using (var entryStream = entry.Open())
		using (StreamWriter writer = new StreamWriter(entryStream, Utf8WithBom))
		{
			writer.Write(sb.ToString());
			Console.WriteLine($"memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");
		}
	}
	Console.WriteLine("end memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");

	Console.WriteLine(path);
}

// result : ![](https://i.imgur.com/q9xYrFs.png)
void Main()
{
	byte[] bytes = Encoding.ASCII.GetBytes("Hello World\n\r");

	Stopwatch sw = new Stopwatch();
	sw.Start();
	Console.WriteLine("start memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB");

	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.zip");
	using (var stream = File.Create(path))
	using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true, Utf8WithBom))
	{
		var entry = archive.CreateEntry("test.txt");
		using (var entryStream = entry.Open())
		{
			for (var i = 0; i <= 10000000; i++)
			{
				entryStream.Write(bytes); //result : ![](https://i.imgur.com/jWEfp6u.png)
				if (i % 1000000 == 0)
					Console.WriteLine($"time.{i} memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");
			}
		}
	}
	Console.WriteLine("end memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");

	Console.WriteLine(path);
}

// You can define other methods, fields, classes and namespaces here
private readonly static UTF8Encoding Utf8WithBom = new System.Text.UTF8Encoding(true);