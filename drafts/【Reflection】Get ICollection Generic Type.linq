<Query Kind="Program">
  <Connection>
    <ID>5fffb9dc-a56f-4ffa-a582-f9da6bc9fdad</ID>
    <Persist>true</Persist>
    <Server>192.168.1.4</Server>
    <SqlSecurity>true</SqlSecurity>
    <UserName>sa</UserName>
    <Password>AQAAANCMnd8BFdERjHoAwE/Cl+sBAAAAyumoRWrbXEqda8ynsoawYAAAAAACAAAAAAAQZgAAAAEAACAAAACumZoBhp4lj0R4mTg98suX0pykwksNIARbRIh49xu5/QAAAAAOgAAAAAIAACAAAADoNABocqodkXYmDtdW0GqBvGuMfAJeL++I3kdCYqM4rxAAAADANn2PCQ6OByhczsa8iMQPQAAAAKx4dlXxPcHN4uDHZRcYbnhkQZ52tjk6YEm+q+GruBVhVPrtz22hjCT4VMaK2N6EtZF2Rfr2P8fUTQH/ZPns5GA=</Password>
    <Database>kn2015</Database>
  </Connection>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Xunit</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Diagnostics</RemoveNamespace>
  <RemoveNamespace>System.IO</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Text</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

//[c# Reflection - Find the Generic Type of a Collection - Stack Overflow](https://stackoverflow.com/questions/2561070/c-sharp-reflection-find-the-generic-type-of-a-collection)
#load "xunit"

#region private::Tests

[Fact]
void ValueGenericTypeTest()
{
	var strings = new List<int>();
	var props = Helpers.GetSubtypeGetProperties(strings);
	Assert.Equal(0,props.Length);
}

public class TestType
{
	public string A { get; set; }
	public string B { get; set; }
}

[Fact()]
void IListUpcastingTest()
{
	IList datas = new List<TestType>();
	var props = Helpers.GetSubtypeGetProperties(datas).ToList();
	Assert.Equal(2, props.Count());
	Assert.Equal("A", props[0].Name);
	Assert.Equal("B", props[1].Name);
}


[Fact()]
void ArrayTest()
{
	ICollection datas = new[] { new { A = "1", B = "2" } };
	var props = Helpers.GetSubtypeGetProperties(datas);
	Assert.Equal(2, props.Count());
}

[Fact()]
void OnlyValidOGenericTypes_Test()
{
	ICollection datas = new[] { new { A = "1", B = "2" } };
	var df = datas.GetType().GetGenericTypeDefinition(); //InvalidOperationException: This operation is only valid on generic types.
}

[Fact()]
void DictionaryTest()
{
	ICollection datas = new[] { new Dictionary<string, object>() { { "A", "A" }, { "B", "B" } } };
	var props = Helpers.GetSubtypeGetProperties(datas);
}
#endregion

internal static class Helpers
{
	public static PropertyInfo[] GetSubtypeGetProperties(ICollection value)
	{
		var collectionType = value.GetType();

		Type gType;
		if (collectionType.IsGenericTypeDefinition || collectionType.IsGenericType)
			gType = collectionType.GetGenericArguments().Single();
		else if (collectionType.IsArray)
			gType = collectionType.GetElementType();
		else
			throw new NotImplementedException($"{collectionType.Name} type not implemented,please issue for me, https://github.com/shps951023/MiniExcel/issues");
		if (typeof(IDictionary).IsAssignableFrom(gType))
			throw new NotImplementedException($"{gType.Name} type not implemented,please issue for me, https://github.com/shps951023/MiniExcel/issues");
		var props = gType.GetProperties(BindingFlags.Public | BindingFlags.Instance);
		return props;
	}
}


void Main()
{
	//RunTests();  // Call RunTests() or press Alt+Shift+T to initiate testing.
}

// You can define other methods, fields, classes and namespaces here

