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
</Query>

void Main()
{
	var st = new Stopwatch();
	st.Start();
	Console.WriteLine($"time : {st.ElapsedMilliseconds} ms");
	Console.WriteLine($"memory usage : {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");

	var index = 0;
	var CellIEnumerableValues = Enumerable.Range(1, 100_000_000)
		.Select(s => "psjdfpsdfjpsdjfpj2p3jp4j23pj4p23j4pj32p4j23p4fspdjfpsdjpfjspdfjpsdfjpsodjfpsodjfposjdpfojsdpfojspdofjspdofjpsdojfposdjfpsdjf")
		//.ToDictionary((s)=>index++,s=>s)
		.ToList()
		//.ToHashSet()
		;
		CellIEnumerableValues.Count();
	Console.WriteLine(GetCount(CellIEnumerableValues));
	Console.WriteLine($"time : {st.ElapsedMilliseconds} ms");
	Console.WriteLine($"memory usage : {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");
}

/*




==== IEnumerable foreach check ====
IEnumerable
	time : 0 ms
	memory usage : 51 MB
	100000000
	time : 1051 ms
	memory usage : 51 MB

ToDictionary
	time : 0 ms
	memory usage : 46 MB
	100000000
	time : 17894 ms
	memory usage : 7371 MB

ToList
	time : 0 ms
	memory usage : 42 MB
	100000000
	time : 1661 ms
	memory usage : 806 MB


==== Auto Check ====
IEnumerable
	time : 0 ms
	memory usage : 49 MB
	100000000
	time : 781 ms
	memory usage : 51 MB
	
ToArray
	time : 0 ms
	memory usage : 49 MB
	100000000
	time : 659 ms
	memory usage : 815 MB

ToDictionary
	time : 0 ms
	memory usage : 46 MB
	ICollection
	100000000
	time : 13225 ms
	memory usage : 7368 MB
	
ToList
	time : 0 ms
	memory usage : 38 MB
	ICollection
	100000000
	time : 850 ms
	memory usage : 802 MB
	
ToHashSet
	time : 0 ms
	memory usage : 39 MB
	IEnumerable
	1
	time : 4295 ms
	memory usage : 39 MB
*/

// You can define other methods, fields, classes and namespaces here
public int GetCount(object value)
{
	{
		var index = 0;
		foreach (var element in value as IEnumerable)
		{
			index++;
		}
		return index;
	}
	throw new Exception();
	var type = value.GetType();
	var prop = type.GetProperty("Count");
	if(prop != null)
	{
		Console.WriteLine("");
		var cnt = prop.GetValue(value);
		return (int) cnt ;
	}
		

	if (value is IEnumerable)
	{
		if(value is ICollection)
		{
			Console.WriteLine("ICollection");
			return (value as ICollection).Count;
		}
		else if (value is IList)
		{
			Console.WriteLine("IList");
			return (value as IList).Count;
		}
		else if (value is Array)
		{
			Console.WriteLine("Array");
			return (value as Array).Length;
		}
		else
		{
			Console.WriteLine("IEnumerable");
			var index = 0;
			foreach (var element in value as IEnumerable)
			{
				index++;
			}
			return index;
		}
	}else{
		throw new Exception("without count");
	}
}