<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>SuperXML</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>Xunit</Namespace>
  <Namespace>SuperXML</Namespace>
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

#load "xunit"
//[beto-rodriguez/SuperXml: A template engine based in AngujarJS markup](https://github.com/beto-rodriguez/SuperXml)
void Main()
{
	Example4();
}

void Example3()
{
	var compiler = new Compiler()
					.AddKey("name", "Excel")
					.AddKey("width", 100)
					.AddKey("height", 500)
					.AddKey("bounds", new[] { 10, 0, 10, 0 })
					.AddKey("elements", new[]
					{
					new { name = "John", age= 10 },
					new { name = "Maria", age= 57 },
					new { name = "Mark", age= 23 },
					new { name = "Edit", age= 82 },
					new { name = "Susan", age= 37 }
					});
	string result = compiler.CompileXml(new StringReader(@"<document>
  <name>my name is {{name}}</name>
  <width>{{width}}</width>
  <height>{{height}}</height>
  <area>{{width*height}}</area>
  <padding>
    <bound sxRepeat=""bound in bounds"">{{bound}}</bound>
  </padding>
  <content>
    <element sxRepeat=""element in elements"" sxIf=""element.age > 25"">
      <name>{{element.name}}</name>
      <age>{{element.age}}</age>
    </element>
  </content> 
</document>"));
	Console.WriteLine(result);
}

void Example1()
{
	var compiler = new Compiler()
					.AddKey("name", "Excel")
					.AddKey("width", 100)
					.AddKey("height", 500)
					.AddKey("bounds", new[] { 10, 0, 10, 0 })
					.AddKey("elements", new[]
					{
					new { name = "John", age= 10 },
					new { name = "Maria", age= 57 },
					new { name = "Mark", age= 23 },
					new { name = "Edit", age= 82 },
					new { name = "Susan", age= 37 }
					});
	string result = compiler.CompileString(@"Hello {{name}}, you are a document with a size of {{width}}x{{height}} and an 
area of {{width*height}}

now here is a list with your bounds:
  <sxRun sxRepeat=""b in bounds"">-value {{$index}}: {{b}}
  </sxRun>

now here you can see a filtered list of classes
  <sxRun sxRepeat=""e in elements"" sxIf=""e.age > 25"">-{{e.name}}, age {{e.age}}
  </sxRun>");
	Console.WriteLine(result);
}

void Example2()
{
	Compiler compiler = new Compiler();

	// 2. Add Elements to your Scope, the first parameter is key, second is value
	//      key:    the 'variable name' for the compiler
	//      value:  the value of the variable in this case the string "world"
	compiler.AddKey("name", "world");

	//3. Call the compile Method and feed the template t get the result
	string result = compiler.CompileString("Hello {{name}}!");
	Console.WriteLine(result);
}

void Example4()
{
	var compiler = new Compiler()
		.AddKey("title", "FooCompany")
		.AddKey("managers", new[] {
			new {name="Jack",department="HR"},
			new {name="Loan",department="IT"}
		})
		.AddKey("employees", new[] {
			new {name="Wade",department="HR"},
			new {name="Felix",department="HR"},
			new {name="Eric",department="IT"},
			new {name="Keaton",department="IT"}
		})
	;

	var result = compiler.CompileXml(new StringReader(template));//a xml string;
	Console.WriteLine(result);
	Console.WriteLine();
	//Assert.Equal(expected, result);
}

const string template = @"<table>
    <tr>
        <td>{{title}}</td>
    </tr>	
	<tr sxRepeat=""manager in managers"">
        <td>{{manager.name}}</td>
		<td>{{manager.department}}</td>
		<td>111</td>
    </tr>		
    <tr>
        <td>Employees</td>
    </tr>
	<tr sxRepeat=""employee in employees"">
        <td>{{employee.name}}</td>
		<td>{{employee.department}}</td>
    </tr>	
</table>";

const string expected = @"<table>
  <tr>
    <td>FooCompany</td>
  </tr>
  <tr>
    <td>Jack</td>
    <td>HR</td>
  </tr>
  <tr>
    <td>Loan</td>
    <td>IT</td>
  </tr>
  <tr>
    <td>Employees</td>
  </tr>
  <tr>
    <td>Wade</td>
    <td>HR</td>
  </tr>
  <tr>
    <td>Felix</td>
    <td>HR</td>
  </tr>
  <tr>
    <td>Eric</td>
    <td>IT</td>
  </tr>
  <tr>
    <td>Keaton</td>
    <td>IT</td>
  </tr>
</table>";

void Example5()
{
	var template = @"
    <table>
        <tr>
            <td>{{title}}</td>
        </tr>	
		<tr>
            <td>{{managers.name}}</td>
			<td>{{managers.department}}</td>
        </tr>		
        <tr>
            <td>Employees</td>
        </tr>
		<tr>
            <td>{{employees.name}}</td>
			<td>{{employees.department}}</td>
        </tr>	
    </table>
	";
	var input = new
	{
		title = "FooCompany",
		managers = new[] {
			new {name="Jack",department="HR"},
			new {name="Loan",department="IT"}
		},
		employees = new[] {
			new {name="Wade",department="HR"},
			new {name="Felix",department="HR"},
			new {name="Eric",department="IT"},
			new {name="Keaton",department="IT"}
		}
	};
	var output = GenerateByTemplate(template, input);
	var expected = @"
    <table>
        <tr>
            <td>FooCompany</td>
        </tr>	
		<tr>
            <td>Jack</td>
			<td>HR</td>
        </tr>	
		<tr>
            <td>Loan</td>
			<td>IT</td>
        </tr>				
        <tr>
            <td>Employees</td>
        </tr>
		<tr>
            <td>Wade</td>
			<td>HR</td>
        </tr>	
		<tr>
            <td>Felix</td>
			<td>HR</td>
        </tr>
		<tr>
            <td>Eric</td>
			<td>IT</td>
        </tr>
		<tr>
            <td>Keaton</td>
			<td>IT</td>
        </tr>		
    </table>	
	";
	//Assert.Equal(expected, output);
}

string GenerateByTemplate(string template, object input)
{
	var compiler = new Compiler();
	{
		var type = input.GetType();
		var props = type.GetProperties();
		foreach (var p in props)
		{
			Console.WriteLine(p.Name);
			Console.WriteLine(p.GetValue(input));
			compiler.AddKey(p.Name, p.GetValue(input));
		}
	}

	var result = compiler.CompileXml(new StringReader(template));//a xml string;
	Console.WriteLine(result);
	return result;
}

// You can define other methods, fields, classes and namespaces here
