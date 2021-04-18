<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.IO.Compression</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
</Query>

void Main()
{
	Test();
	Parallel.ForEach(Enumerable.Range(1, 5), s =>
	{
		Console.WriteLine(s);
		try
		{
			Test();
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex);
			throw;
		}

	});
}

// You can define other methods, fields, classes and namespaces here
void Test()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
	var templatePath = @"D:\git\MiniExcel\samples\xlsx\TestTemplateGithubProjects.xlsx";
	var projects = new[]
	{
		new {Name = "MiniExcel",Link="https://github.com/shps951023/MiniExcel",Star=146, CreateTime=new DateTime(2021,03,01)},
		new {Name = "HtmlTableHelper",Link="https://github.com/shps951023/HtmlTableHelper",Star=16, CreateTime=new DateTime(2020,02,01)},
		new {Name = "PocoClassGenerator",Link="https://github.com/shps951023/PocoClassGenerator",Star=16, CreateTime=new DateTime(2019,03,17)}
	};
	var value = new
	{
		User = "ITWeiHan",
		Projects = projects,
		TotalStar = projects.Sum(s => s.Star)
	};
	using (var stream = File.Create(path))
	{
		stream.SaveAsByTemplate(templatePath, value);
		Thread.Sleep(500);
	}
	//MiniExcel.SaveAsByTemplate(path, templatePath, value);
	Console.WriteLine(path);
}