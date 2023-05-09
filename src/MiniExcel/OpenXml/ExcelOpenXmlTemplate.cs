
namespace MiniExcelLibs.OpenXml
{
    using MiniExcelLibs.Utils;
    using MiniExcelLibs.Zip;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices.ComTypes;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Xml;

    internal partial class ExcelOpenXmlTemplate : IExcelTemplate, IExcelTemplateAsync
    {
        private static readonly XmlNamespaceManager _ns;
        private static readonly Regex _isExpressionRegex;
        static ExcelOpenXmlTemplate()
        {
            _isExpressionRegex = new Regex("(?<={{).*?(?=}})");
            _ns = new XmlNamespaceManager(new NameTable());
            _ns.AddNamespace("x", Config.SpreadsheetmlXmlns);
        }

        private readonly Stream _stream;
        private readonly OpenXmlConfiguration _configuration;
        private readonly StringBuilder _calcChainContent = new StringBuilder();

        public ExcelOpenXmlTemplate(Stream stream, IConfiguration configuration)
        {
            _stream = stream;
            _configuration = (OpenXmlConfiguration)configuration ?? OpenXmlConfiguration.DefaultConfig;
        }

        public void SaveAsByTemplate(string templatePath, object value)
        {
            using (var stream = FileHelper.OpenSharedRead(templatePath))
                SaveAsByTemplateImpl(stream, value);
        }
        public void SaveAsByTemplate(byte[] templateBtyes, object value)
        {
            using (Stream stream = new MemoryStream(templateBtyes))
                SaveAsByTemplateImpl(stream, value);
        }

        public void SaveAsByTemplateImpl(Stream templateStream, object value)
        {
            //only support xlsx         
            Dictionary<string, object> values = null;
            if (value is Dictionary<string, object>)
            {
                values = value as Dictionary<string, object>;
                foreach (var key in values.Keys)
                {
                    var v = values[key];
                    if (v is IDataReader)
                    {
                        values[key] = TypeHelper.ConvertToEnumerableDictionary(v as IDataReader).ToList();
                    }
                }
            }
            else
            {
                var type = value.GetType();
                var props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
                values = new Dictionary<string, object>();
                foreach (var p in props)
                {
                    values.Add(p.Name, p.GetValue(value));
                }

                var fields = type.GetFields(BindingFlags.Public | BindingFlags.Instance);
                foreach (var f in fields)
                {
                    if (!values.ContainsKey(f.Name))
                    { values.Add(f.Name, f.GetValue(value)); }
                }
            }

            {
                templateStream.CopyTo(_stream);

                var reader = new ExcelOpenXmlSheetReader(_stream, null);
                var _archive = new ExcelOpenXmlZip(_stream, mode: ZipArchiveMode.Update, true, Encoding.UTF8);
                {
                    //read sharedString
                    var sharedStrings = reader._sharedStrings;
                    StringBuilder calcSheetContent = new StringBuilder();

                    //read all xlsx sheets
                    var sheets = _archive.zipFile.Entries.Where(w => w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                        || w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                    ).ToList();

                    int sheetIdx = 0;
                    foreach (var sheet in sheets)
                    {
                        this.XRowInfos = new List<XRowInfo>(); //every time need to use new XRowInfos or it'll cause duplicate problem: https://user-images.githubusercontent.com/12729184/115003101-0fcab700-9ed8-11eb-9151-ca4d7b86d59e.png
                        this.XMergeCellInfos = new Dictionary<string, XMergeCell>();
                        this.NewXMergeCellInfos = new List<XMergeCell>();

                        var sheetStream = sheet.Open();
                        var fullName = sheet.FullName;

                        ZipArchiveEntry entry = _archive.zipFile.CreateEntry(fullName);
                        using (var zipStream = entry.Open())
                        {
                            GenerateSheetXmlImpl(sheet, zipStream, sheetStream, values, sharedStrings, false);
                            //doc.Save(zipStream); //don't do it because : ![image](https://user-images.githubusercontent.com/12729184/114361127-61a5d100-9ba8-11eb-9bb9-34f076ee28a2.png)
                        }

                        // disposing writer disposes streams as well. reopen the entry to read and parse calc functions
                        using (var filledStream = entry.Open())
                        {
                            sheetIdx++; 
                            var filledDoc = new XmlDocument();
                            filledDoc.Load(filledStream);
                            var filledSheetData = filledDoc.SelectSingleNode("/x:worksheet/x:sheetData", _ns);
                            _calcChainContent.Append(CalcChainHelper.GetCalcChainContentFromSheet(filledSheetData, _ns, sheetIdx));
                        }
                    }

                    var calcChain = _archive.zipFile.Entries.FirstOrDefault( e => e.FullName.Contains("xl/calcChain.xml"));
                    if (calcChain != null)
                    {
                        string calcChainPathname = calcChain.FullName;
                        calcChain.Delete();
                        var calcChainEntry =  _archive.zipFile.CreateEntry(calcChainPathname);
                        using (var calcChainStream = calcChainEntry.Open())
                        {
                            CalcChainHelper.GenerateCalcChainSheet(calcChainStream, _calcChainContent.ToString());
                        }
                    }
                    
                }

                _archive.zipFile.Dispose();
            }
        }

        public Task SaveAsByTemplateAsync(string templatePath, object value, CancellationToken cancellationToken = default(CancellationToken))
        {
            return Task.Run(() => SaveAsByTemplate(templatePath, value), cancellationToken);
        }

        public Task SaveAsByTemplateAsync(byte[] templateBtyes, object value, CancellationToken cancellationToken = default(CancellationToken))
        {
            return Task.Run(() => SaveAsByTemplate(templateBtyes, value), cancellationToken);
        }
    }
}
