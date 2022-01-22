using MiniExcelLibs.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MiniExcelLibs.Csv
{
    internal class CsvWriter : IExcelWriter
    {
        private readonly Stream _stream;
        private readonly CsvConfiguration _configuration;
        private readonly bool _printHeader;
        private readonly object _value;

        public CsvWriter(Stream stream, object value, IConfiguration configuration, bool printHeader)
        {
            this._stream = stream;
            this._configuration = configuration == null ? CsvConfiguration.DefaultConfiguration : (CsvConfiguration)configuration;
            this._printHeader = printHeader;
            this._value = value;
        }

        public void SaveAs()
        {
            var seperator = _configuration.Seperator.ToString();
            var newLine = _configuration.NewLine;

            using (StreamWriter writer = _configuration.StreamWriterFunc(_stream))
            {
                if (_value == null)
                {
                    writer.Write("");
                    return;
                }

                var type = _value.GetType();
                Type genericType = null;

                if (_value is IDataReader)
                {
                    GenerateSheetByIDataReader(_value, seperator, newLine, writer);
                }
                else if (_value is IEnumerable)
                {
                    var values = _value as IEnumerable;
                    List<object> keys = new List<object>();
                    List<ExcelCustomPropertyInfo> props = null;
                    string mode = null;

                    // check mode
                    {
                        foreach (var item in values) //TODO: need to optimize
                        {
                            if (item != null && mode == null)
                            {
                                if (item is IDictionary<string, object>)
                                {
                                    var item2 = item as IDictionary<string, object>;
                                    mode = "IDictionary<string, object>";
                                    foreach (var key in item2.Keys)
                                        keys.Add(key);
                                }
                                else if (item is IDictionary)
                                {
                                    var item2 = item as IDictionary;
                                    mode = "IDictionary";
                                    foreach (var key in item2.Keys)
                                        keys.Add(key);
                                }
                                else
                                {
                                    mode = "Properties";
                                    genericType = item.GetType();
                                    props = CustomPropertyHelper.GetSaveAsProperties(genericType);
                                }

                                break;
                            }
                        }
                    }

                    //if(mode == null)
                    //    throw new NotImplementedException($"Type {type?.Name} & genericType {genericType?.Name} not Implemented. please issue for me.");

                    if (keys.Count == 0 && props == null)
                    {
                        writer.Write(newLine);
                        return;
                    }

                    if (this._printHeader)
                    {
                        if (props != null)
                        {
                            writer.Write(string.Join(seperator, props.Select(s => CsvHelpers.ConvertToCsvValue(s?.ExcelColumnName))));
                            writer.Write(newLine);
                        }
                        else if (keys.Count > 0)
                        {
                            writer.Write(string.Join(seperator, keys));
                            writer.Write(newLine);
                        }
                        else
                        {
                            throw new InvalidOperationException("Please issue for me.");
                        }
                    }

                    if (mode == "IDictionary<string, object>") //Dapper Row
                        GenerateSheetByDapperRow(writer, _value as IEnumerable, keys.Cast<string>().ToList(), seperator, newLine);
                    else if (mode == "IDictionary") //IDictionary
                        GenerateSheetByIDictionary(writer, _value as IEnumerable, keys, seperator, newLine);
                    else if (mode == "Properties")
                        GenerateSheetByProperties(writer, _value as IEnumerable, props, seperator, newLine);
                    else
                        throw new NotImplementedException($"Type {type?.Name} & genericType {genericType?.Name} not Implemented. please issue for me.");
                }
                else if (_value is DataTable)
                {
                    GenerateSheetByDataTable(writer, _value as DataTable, seperator, newLine);
                }
                else
                {
                    throw new NotImplementedException($"Type {type?.Name} & genericType {genericType?.Name} not Implemented. please issue for me.");
                }
            }
        }

        public Task SaveAsAsync()
        {
            return Task.Run(() => SaveAs());
        }

        private void GenerateSheetByIDataReader(object value, string seperator, string newLine, StreamWriter writer)
        {
            var reader = (IDataReader)value;

            int fieldCount = reader.FieldCount;
            if (fieldCount == 0)
                throw new InvalidDataException("fieldCount is 0");

            if (this._printHeader)
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    var columnName = reader.GetName(i);

                    if (i != 0)
                        writer.Write(seperator);
                    writer.Write(CsvHelpers.ConvertToCsvValue(columnName?.ToCsvString(null)));
                }
                writer.Write(newLine);
            }

            while (reader.Read())
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    var cellValue = reader.GetValue(i);
                    if (i != 0)
                        writer.Write(seperator);
                    writer.Write(CsvHelpers.ConvertToCsvValue(cellValue?.ToCsvString(null)));
                }
                writer.Write(newLine);
            }
        }

        private void GenerateSheetByDataTable(StreamWriter writer, DataTable dt, string seperator, string newLine)
        {
            if (_printHeader)
            {
                writer.Write(string.Join(seperator, dt.Columns.Cast<DataColumn>().Select(s => s.Caption ?? s.ColumnName)));
                writer.Write(newLine);
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var first = true;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    var cellValue = CsvHelpers.ConvertToCsvValue(dt.Rows[i][j]?.ToCsvString(null));
                    if (!first)
                        writer.Write(seperator);
                    writer.Write(cellValue);
                    first = false;
                }
                writer.Write(newLine);
            }
        }

        private void GenerateSheetByProperties(StreamWriter writer, IEnumerable value, List<ExcelCustomPropertyInfo> props, string seperator, string newLine)
        {
            foreach (var v in value)
            {
                var values = props.Select(s => CsvHelpers.ConvertToCsvValue(s?.Property.GetValue(v)?.ToCsvString(s)));
                writer.Write(string.Join(seperator, values));
                writer.Write(newLine);
            }
        }

        private void GenerateSheetByIDictionary(StreamWriter writer, IEnumerable value, List<object> keys, string seperator, string newLine)
        {
            foreach (IDictionary v in value)
            {
                var values = keys.Select(key => CsvHelpers.ConvertToCsvValue(v[key]?.ToCsvString(null)));
                writer.Write(string.Join(seperator, values));
                writer.Write(newLine);
            }
        }

        private void GenerateSheetByDapperRow(StreamWriter writer, IEnumerable value, List<string> keys, string seperator, string newLine)
        {
            foreach (IDictionary<string, object> v in value)
            {
                var values = keys.Select(key => CsvHelpers.ConvertToCsvValue(v[key]?.ToCsvString(null)));
                writer.Write(string.Join(seperator, values));
                writer.Write(newLine);
            }
        }
    }

    internal static class CsvValueTostringHelper
    {
        public static string ToCsvString(this object value, ExcelCustomPropertyInfo p)
        {
            if (value == null)
                return "";

            Type type = null;
            if (p == null)
            {
                type = value.GetType();
                type = Nullable.GetUnderlyingType(type) ?? type;
            }
            else
            {
                type = p.ExcludeNullableType; //sometime it doesn't need to re-get type like prop
            }

            if (type == typeof(DateTime))
            {
                if (p == null || p.ExcelFormat == null)
                {
                    return ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss");
                }
                else
                {
                    return ((DateTime)value).ToString(p.ExcelFormat);
                }
            }

            return value.ToString();
        }
    }
}
