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
	Console.WriteLine("====PassingRefByVal====");
	PassingRefByVal.Main();
	Console.WriteLine("====PassingRefByRef====");
	PassingRefByRef.Main();
}

// You can define other methods, fields, classes and namespaces here
class PassingRefByRef
{
	static void Change(ref int[] pArray)
	{
		// Both of the following changes will affect the original variables:
		pArray[0] = 888;
		pArray = new int[5] { -3, -1, -2, -3, -4 };
		System.Console.WriteLine("Inside the method, the first element is: {0}", pArray[0]);
	}

	public static void Main()
	{
		
		int[] arr = { 1, 4, 5 };
		System.Console.WriteLine("Inside Main, before calling the method, the first element is: {0}", arr[0]);

		Change(ref arr);
		System.Console.WriteLine("Inside Main, after calling the method, the first element is: {0}", arr[0]);
	}
}

class PassingRefByVal
{
	static void Change(int[] pArray)
	{
		pArray[0] = 888;  // This change affects the original element.
		pArray = new int[5] { -3, -1, -2, -3, -4 };   // This change is local.
		System.Console.WriteLine("Inside the method, the first element is: {0}", pArray[0]);
	}

	public static void Main()
	{
		int[] arr = { 1, 4, 5 };
		System.Console.WriteLine("Inside Main, before calling the method, the first element is: {0}", arr[0]);

		Change(arr);
		System.Console.WriteLine("Inside Main, after calling the method, the first element is: {0}", arr[0]);
	}
}