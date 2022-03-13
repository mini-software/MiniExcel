using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    public class MiniExcelDynamic
    {
        FileStream stream;
        ExcelOpenXmlZip archive;
        ExcelOpenXmlSheetReader excelOpenXmlSheetReader;
        public delegate void MyDelegate(string str);//用于输出异常信息
        public MyDelegate ShowMsgHandler;

        public bool Open(string path, IConfiguration configuration = null)
        {
            bool result = false;
            if(!File.Exists(path))return result;
            try
            {
                stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                archive = new ExcelOpenXmlZip(stream);
                if (ExcelTypeHelper.GetExcelType(stream, ExcelType.UNKNOWN) != ExcelType.XLSX) return false;
                excelOpenXmlSheetReader = new ExcelOpenXmlSheetReader(stream, configuration);
                result = true;
            }
            catch(Exception ex)
            {
                result = false;
                showMsg(ex.Message);
            }
            return result;
        }

        public void Close()
        {
            stream?.Close();
            archive?.Dispose();
            //excelOpenXmlSheetReader?.Dispose();
        }

        public IEnumerable<T> Query<T>(string sheetName = null, string startCell = "A1") where T : class, new()
        {
            return excelOpenXmlSheetReader.Query<T>(sheetName, startCell);
        }
        public IEnumerable<dynamic> Query(bool useHeaderRow = false, string sheetName = null, string startCell = "A1")
        {
            return excelOpenXmlSheetReader.Query(useHeaderRow, sheetName, startCell);
        }
        public DataTable QueryAsDataTable(bool useHeaderRow = true, string sheetName = null, string startCell = "A1")
        {
            if (sheetName == null) /*Issue #279*/
                sheetName = stream.GetSheetNames().First();

            var dt = new DataTable(sheetName);
            var first = true;
            var rows = excelOpenXmlSheetReader.Query(useHeaderRow, sheetName, startCell);

            var keys = new List<string>();
            foreach (IDictionary<string, object> row in rows)
            {
                if (first)
                {
                    foreach (var key in row.Keys)
                    {
                        if (!string.IsNullOrEmpty(key)) // avoid #298 : Column '' does not belong to table
                        {
                            var column = new DataColumn(key, typeof(object)) { Caption = key };
                            dt.Columns.Add(column);
                            keys.Add(key);
                        }
                    }

                    dt.BeginLoadData();
                    first = false;
                }

                var newRow = dt.NewRow();
                foreach (var key in keys)
                {
                    newRow[key] = row[key]; //TODO: optimize not using string key
                }

                dt.Rows.Add(newRow);
            }

            dt.EndLoadData();
            return dt;
        }
        public List<string> GetSheetNames()
        {
            List<string> sheetNames=new List<string>();
            if (stream == null)return sheetNames;
            sheetNames = excelOpenXmlSheetReader.GetWorkbookRels(archive.entries).Select(s => s.Name).ToList();
            
            return sheetNames;
        }

        private void showMsg(string msg)
        {
            if(ShowMsgHandler!=null) ShowMsgHandler(msg);
        }
    }
}
