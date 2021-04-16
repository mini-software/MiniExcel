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
	Console.WriteLine($"start memory usage : {GetCurrentMemoryUsage()} MB");
	
	var sb = new StringBuilder();
	for (int i = 0; i < 10000000; i++)
	{
		sb.AppendLine(RemoveNamespace($"fsdfsdpfjsdpfjpdsjp32j4p23j4p23423423jpjpjpjpojsdfpjsdfdsfsdfsd  xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""));
	}
	
	Console.WriteLine($"time : {st.ElapsedMilliseconds} ms");
	Console.WriteLine($"end memory usage : {GetCurrentMemoryUsage()} MB");
}

public decimal GetCurrentMemoryUsage() => Process.GetCurrentProcess().WorkingSet64/(1024*1024);

/*
Running 10,000,000 times Reuslt
not replace:
	time : 0 ms
	start memory usage : 46 MB
	time : 3345 ms
	end memory usage : 4025 MB
	
replace 
	time : 0 ms
	start memory usage : 46 MB
	time : 5128 ms
	end memory usage : 1353 MB
*/

// You can define other methods, fields, classes and namespaces here
private static string RemoveNamespace(string xml)
{
	//TODO: need to optimize
	return xml
		.Replace("xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"", "").Replace("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "")
	;
}