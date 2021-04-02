<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
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

void Main()
{
	var input = new Demo
	{
		Title="Basic Demo",
		Users = new User[] {
			new User{ID=Guid.NewGuid(),Name="Jack",Age=25,InDate=new DateTime(2021,3,1),VIP=true,Points=new Decimal(1234.55)},
			new User{ID=Guid.NewGuid(),Name="Lisa",Age=44,InDate=new DateTime(2021,2,14),VIP=false,Points=new Decimal(5741.201)},
		}
	};

	var type = input.GetType();
	var level1props = type.GetProperties();
	foreach (var s in level1props)
	{
		//PropertyType vs ReflectedType vs DeclaringType vs MemberType
		var pt = s.PropertyType;  //UserQuery+User[]
		var mt = s.MemberType;
		var rt = s.ReflectedType; //UserQuery+Demo
		var dt = s.DeclaringType; //UserQuery+Demo
		var b1 = typeof(IEnumerable).IsAssignableFrom(pt); //true 
		var b3 = typeof(IEnumerable).IsAssignableFrom(rt);  //false
		var b4 = typeof(IEnumerable).IsAssignableFrom(dt); //false
	}
}

// You can define other methods, fields, classes and namespaces here
public class Demo
{
	public string Title { get; set; }
	public User[] Users { get; set; }	
}
public class User
{
	public Guid ID { get; set; }
	public string Name { get; set; }
	public int Age { get; set; }
	public DateTime InDate { get; set; }
	public bool VIP { get; set; }
	public decimal Points { get; set; }
}