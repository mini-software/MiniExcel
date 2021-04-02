<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.ComponentModel.DataAnnotations</Namespace>
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

// [Using Data Annotations to Validate Models in .NET]( https://www.c-sharpcorner.com/UploadFile/20c06b/using-data-annotations-to-validate-models-in-net/ )

void Main()
{
	Test1();
}

void Test1()
{
	{


		var invalidGame = new Game
		{
			Name = "My name is way over 20 characters",
			Price = 300,
		};

		ICollection<ValidationResult> results = null;
		if (!Validate(invalidGame, out results))
			Console.WriteLine(String.Join("\n", results.Select(o => o.ErrorMessage)));
		else
			Console.WriteLine("I'm a valid object!");
	}
	{
		var validGame = new Game
		{
			Name = "Magicka",
			Price = 5,
		};
		ICollection<ValidationResult> results = null;
		if (!Validate(validGame, out results))
			Console.WriteLine(String.Join("\n", results.Select(o => o.ErrorMessage)));
		else
			Console.WriteLine("I'm a valid object!");
	}
}

static bool Validate<T>(T obj, out ICollection<ValidationResult> results)
{
	results = new List<ValidationResult>();

	return Validator.TryValidateObject(obj, new ValidationContext(obj), results, true);
}

// You can define other methods, fields, classes and namespaces here
public class Game
{
	[Required]
	[StringLength(20)]
	public string Name { get; set; }

	[Range(0, 100)]
	public decimal Price { get; set; }
}