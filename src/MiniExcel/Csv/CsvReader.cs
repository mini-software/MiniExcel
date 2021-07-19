using MiniExcelLibs.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MiniExcelLibs.Csv
{
    internal class CsvReader : IExcelReader , IExcelReaderAsync
    {
        private Stream _stream;
        public CsvReader(Stream stream)
        {
            this._stream = stream;
        }
        public IEnumerable<IDictionary<string, object>> Query(bool useHeaderRow, string sheetName, string startCell, IConfiguration configuration)
        {
            if (startCell != "A1")
                throw new NotImplementedException("CSV not Implement startCell");
            var cf = configuration == null ? CsvConfiguration.DefaultConfiguration : (CsvConfiguration)configuration;

            using (var reader = cf.StreamReaderFunc(_stream))
            {
                var row = string.Empty;
                string[] read;
                var firstRow = true;
                Dictionary<int, string> headRows = new Dictionary<int, string>();
                while ((row = reader.ReadLine()) != null)
                {
                    read = Split(cf, row);

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

                        var cell = CustomPropertyHelper.GetEmptyExpandoObject(headRows);
                        for (int i = 0; i <= read.Length - 1; i++)
                            cell[headRows[i]] = read[i];

                        yield return cell;
                        continue;
                    }


                    //body
                    {
                        var cell = CustomPropertyHelper.GetEmptyExpandoObject(read.Length - 1, 0);
                        for (int i = 0; i <= read.Length - 1; i++)
                            cell[ColumnHelper.GetAlphabetColumnName(i)] = read[i];
                        yield return cell;
                    }
                }
            }
        }

        public IEnumerable<T> Query<T>(string sheetName, string startCell, IConfiguration configuration) where T : class, new()
        {
            var cf = configuration == null ? CsvConfiguration.DefaultConfiguration : (CsvConfiguration)configuration;

            var type = typeof(T);

            Dictionary<int, ExcelCustomPropertyInfo> idxProps = new Dictionary<int, ExcelCustomPropertyInfo>();
            using (var reader = cf.StreamReaderFunc(_stream))
            {
                var row = string.Empty;
                string[] read;

                //header
                {
                    row = reader.ReadLine();
                    read = Split(cf, row);

                    var props = CustomPropertyHelper.GetExcelCustomPropertyInfos(type, read);
                    var index = 0;
                    foreach (var v in read)
                    {
                        var p = props.SingleOrDefault(w => w.ExcelColumnName == v);
                        if (p != null)
                            idxProps.Add(index, p);
                        index++;
                    }
                }
                {
                    while ((row = reader.ReadLine()) != null)
                    {
                        read = Split(cf, row);

                        //body
                        {
                            var v = new T();

                            var rowIndex = 0; //TODO: rowindex = startcell rowindex
                            foreach (var p in idxProps)
                            {
                                var pInfo = p.Value;

                                {

                                    object newV = null;
                                    object itemValue = read[p.Key];

                                    if (itemValue == null)
                                        continue;

                                    newV = TypeHelper.TypeMapping(v, pInfo, newV, itemValue, rowIndex, startCell);
                                }
                            }

                            rowIndex++;
                            yield return v;
                        }
                    }

                }
            }
        }

        private static string[] Split(CsvConfiguration cf, string row)
        {
            return Regex.Split(row, $"[\t{cf.Seperator}](?=(?:[^\"]|\"[^\"]*\")*$)")
                .Select(s => Regex.Replace(s.Replace("\"\"", "\""), "^\"|\"$", "")).ToArray();
            //this code from S.O : https://stackoverflow.com/a/11365961/9131476
        }

        public Task<IEnumerable<IDictionary<string, object>>> QueryAsync(bool UseHeaderRow, string sheetName, string startCell, IConfiguration configuration)
        {
            return Task.Run(() => Query(UseHeaderRow, sheetName, startCell, configuration));
        }

        public Task<IEnumerable<T>> QueryAsync<T>(string sheetName, string startCell, IConfiguration configuration) where T : class, new()
        {
            return Task.Run(() => Query<T>(sheetName, startCell, configuration));
        }
    }
}
