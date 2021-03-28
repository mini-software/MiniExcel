using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using MiniExcelLibs;

namespace MiniExcelLibs.Benchmarks
{
    class Program
    {
        static void Main(string[] args)
        {
		  var values = Enumerable.Range(1, 1048575).Select((s, index) => new { index, value = Guid.NewGuid() });
		  var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
		  using (var stream = File.Create(path))
			 stream.SaveAs(values);
	   }
    }
}
