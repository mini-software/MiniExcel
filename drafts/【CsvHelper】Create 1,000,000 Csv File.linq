<Query Kind="Program">
  <NuGetReference>CsvHelper</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.IO.Compression</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>CsvHelper</Namespace>
  <Namespace>System.Globalization</Namespace>
</Query>

void Main()
{
	var records = Enumerable.Range(1,1000000).Select((s,idx)=>new  { Id = idx , Text = "Hello World" });
	var path =  Path.Combine(Path.GetTempPath(),$"{Guid.NewGuid().ToString()}.csv");
	using (var writer = new StreamWriter(path))
	using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
	{
		csv.WriteRecords(records);
	}
	Console.WriteLine(path);
}

