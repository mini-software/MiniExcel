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
  <Namespace>Xunit</Namespace>
</Query>

#load "xunit"
void Main(){
	Generate("\"\"\""); //result : """"""""
	Generate(","); //result : ","
	Generate("\t"); //result : \t,\t
	Generate(" "); //result : " "," "
	//Generate(Environment.NewLine); //result : " "," "
}

void Generate(string value)
{
	var records = Enumerable.Range(1, 1).Select((s, idx) => new { v1 = value,v2=value });
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.csv");
	using (var writer = new StreamWriter(path))
	using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
	{
		csv.WriteRecords(records);
	}
	Console.WriteLine(path);
}