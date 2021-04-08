<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.IO.Compression</Namespace>
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
	Test1();
	Test2();
}

void Test2()
{
	Console.WriteLine("==== using ZipArchive dispose ====");
	Stopwatch sw = new Stopwatch();
	sw.Start();
	var path = @"D:\git\MiniExcel\samples\xlsx\TestMultiSheet.xlsx";


	Console.WriteLine("start memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");

	for (int i = 0; i < 10000; i++)
	{
		using (var stream = File.OpenRead(path))
		using (var zip = new ZipArchive(stream))
		{
			var reader = new ReaderWithZipArchiveDispose();
			reader.GetXmlContent(zip, "xl/worksheets/sheet1.xml").ToList();
			reader.GetXmlContent(zip, "xl/worksheets/sheet2.xml").ToList();
			reader.GetXmlContent(zip, "xl/worksheets/sheet3.xml").ToList();

			if (i % 5000 == 0)
				Console.WriteLine($"no.{i} memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");
		}
	}
	Console.WriteLine("end memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");
}

void Test1()
{
	Console.WriteLine("==== wihout using ZipArchive dispose ====");
	Stopwatch sw = new Stopwatch();
	sw.Start();
	var path = @"D:\git\MiniExcel\samples\xlsx\TestMultiSheet.xlsx";


	Console.WriteLine("start memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");

	for (int i = 0; i < 10000; i++)
	{
		using (var stream = File.OpenRead(path))
		{
			var reader = new ReaderWithoutZipArchiveDispose(stream);
			reader.GetXmlContent(stream, "xl/worksheets/sheet1.xml").ToList();
			reader.GetXmlContent(stream, "xl/worksheets/sheet2.xml").ToList();
			reader.GetXmlContent(stream, "xl/worksheets/sheet3.xml").ToList();

			if (i % 5000 == 0)
				Console.WriteLine($"no.{i} memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");
		}
	}
	Console.WriteLine("end memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");
}

// You can define other methods, fields, classes and namespaces here
public class ReaderWithoutZipArchiveDispose
{
	private ZipArchive _zip;
	public ReaderWithoutZipArchiveDispose(Stream stream)
	{
		_zip = new ZipArchive(stream);
	}

	public string GetXmlContent(Stream stream, string path)
	{
		var c = _zip.Entries.Single(_ => _.FullName.Contains(path));
		using (var s = c.Open())
		using (var sr = new StreamReader(s))
			return sr.ReadToEnd();
	}
}

public class ReaderWithZipArchiveDispose
{
	public string GetXmlContent(ZipArchive _zip, string path)
	{
		var c = _zip.Entries.Single(_ => _.FullName.Contains(path));
		using (var s = c.Open())
		using (var sr = new StreamReader(s))
			return sr.ReadToEnd();
	}
}

