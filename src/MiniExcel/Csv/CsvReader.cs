using MiniExcelLibs.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MiniExcelLibs.Csv
{
    public class CsvReader
    {
	   internal IEnumerable<IDictionary<string, object>> Query(string path, bool useHeaderRow, CsvConfiguration configuration)
	   {
		  if (configuration == null)
			 configuration = CsvConfiguration.GetDefaultConfiguration();
		  using (var stream = File.OpenRead(path))
		  //note: why duplicate code can see #124 issue
		  using (var reader = configuration.GetStreamReaderFunc(stream))
		  {
			 char[] seperators = { configuration.Seperator };

			 var row = string.Empty;
			 string[] read;
			 var firstRow = true;
			 Dictionary<int, string> headRows = new Dictionary<int, string>();
			 while ((row = reader.ReadLine()) != null)
			 {
				read = row.Split(seperators, StringSplitOptions.None);

				//header
				if (useHeaderRow)
				{
				    if (firstRow)
				    {
					   firstRow = false;
					   for (int i = 0; i <= read.Length - 1; i++)
						  headRows.Add(i, read[i]);
					   continue;
				    }

				    var cell = Helpers.GetEmptyExpandoObject(headRows);
				    for (int i = 0; i <= read.Length - 1; i++)
					   cell[headRows[i]] = read[i];

				    yield return cell;
				    continue;
				}


				//body
				{
				    var cell = Helpers.GetEmptyExpandoObject(read.Length - 1);
				    for (int i = 0; i <= read.Length - 1; i++)
					   cell[Helpers.GetAlphabetColumnName(i)] = read[i];
				    yield return cell;
				}
			 }
		  }
	   }

	   internal IEnumerable<IDictionary<string, object>> Query(Stream stream, bool useHeaderRow, CsvConfiguration configuration)
	   {
		  if (configuration == null)
			 configuration = CsvConfiguration.GetDefaultConfiguration();
		  //note: why duplicate code can see #124 issue
		  using (var reader = configuration.GetStreamReaderFunc(stream))
		  {
			 char[] seperators = { configuration.Seperator };

			 var row = string.Empty;
			 string[] read;
			 var firstRow = true;
			 Dictionary<int, string> headRows = new Dictionary<int, string>();
			 while ((row = reader.ReadLine()) != null)
			 {
				read = row.Split(seperators, StringSplitOptions.None);

				//header
				if (useHeaderRow)
				{
				    if (firstRow)
				    {
					   firstRow = false;
					   for (int i = 0; i <= read.Length - 1; i++)
						  headRows.Add(i, read[i]);
					   continue;
				    }

				    var cell = Helpers.GetEmptyExpandoObject(headRows);
				    for (int i = 0; i <= read.Length - 1; i++)
					   cell[headRows[i]] = read[i];

				    yield return cell;
				    continue;
				}


				//body
				{
				    var cell = Helpers.GetEmptyExpandoObject(read.Length - 1);
				    for (int i = 0; i <= read.Length - 1; i++)
					   cell[Helpers.GetAlphabetColumnName(i)] = read[i];
				    yield return cell;
				}
			 }
		  }
	   }
    }
}
