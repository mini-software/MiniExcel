using MiniExcelLibs.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace MiniExcelLibs.Csv
{
    internal class CsvReader : IExcelReader
    {
        private Stream _stream;
        public CsvReader(Stream stream)
        {
            this._stream = stream;
        }
        public IEnumerable<IDictionary<string, object>> Query(bool useHeaderRow, string sheetName, IConfiguration configuration)
        {
            var cf = configuration == null ? CsvConfiguration.DefaultConfiguration : (CsvConfiguration)configuration;

            using (var reader = cf.GetStreamReaderFunc(_stream))
            {
                char[] seperators = { cf.Seperator };

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

        public IEnumerable<T> Query<T>(string sheetName, IConfiguration configuration) where T : class, new()
        {
            var cf = configuration == null ? CsvConfiguration.DefaultConfiguration : (CsvConfiguration)configuration;

            var type = typeof(T);

            Dictionary<int, PropertyInfo> idxProps = new Dictionary<int, PropertyInfo>();
            using (var reader = cf.GetStreamReaderFunc(_stream))
            {
                char[] seperators = { cf.Seperator };

                var row = string.Empty;
                string[] read;

                //header
                {
                    row = reader.ReadLine();
                    read = row.Split(seperators, StringSplitOptions.None);
                    var props = Helpers.GetExcelCustomPropertyInfos(type, read);
                    var index = 0;
                    foreach (var v in read)
                    {
                        var p = props.SingleOrDefault(w => w.ExcelColumnName == v);
                        if (p != null)
                            idxProps.Add(index, p.Property);
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
                            {
                                if (p.Value.PropertyType.IsEnum)
                                {
                                    var newV = Enum.Parse(p.Value.PropertyType, read[p.Key], true);
                                    p.Value.SetValue(cell, newV);
                                }
                                else
                                {
                                    p.Value.SetValue(cell, read[p.Key]);
                                }
                            }

                            yield return cell;
                        }
                    }

                }
            }
        }
    }
}
