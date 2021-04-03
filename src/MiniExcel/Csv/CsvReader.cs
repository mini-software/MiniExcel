using MiniExcelLibs.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MiniExcelLibs.Csv
{
    internal class CsvReader : IExcelReader
    {
        public IEnumerable<IDictionary<string, object>> Query(Stream stream, bool useHeaderRow)
        {
          
            var configuration = new CsvConfiguration();
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

        public IEnumerable<T> Query<T>(Stream stream) where T : class, new()
        {
            var type = typeof(T);
            var props = Helpers.GetProperties(type);
            Dictionary<int, PropertyInfo> idxProps = new Dictionary<int, PropertyInfo>();
            var configuration = new CsvConfiguration();
            using (var reader = configuration.GetStreamReaderFunc(stream))
            {
                char[] seperators = { configuration.Seperator };

                var row = string.Empty;
                string[] read;

                //header
                {
                    row = reader.ReadLine();
                    read = row.Split(seperators, StringSplitOptions.None);
                    var index = 0;
                    foreach (var v in read)
                    {
                        var p = props.SingleOrDefault(w => w.Name == v);
                        if (p != null)
                            idxProps.Add(index,p);
                        index++;
                    }
                }
                {
                    while ((row = reader.ReadLine()) != null)
                    {
                        read = row.Split(seperators, StringSplitOptions.None);

                        //body
                        {
                            var cell = new T();
                            foreach (var p in idxProps)
                                p.Value.SetValue(cell, read[p.Key]);
                            yield return cell;
                        }
                    }

                }
            }
        }
    }
}
