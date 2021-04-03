using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace MiniExcelLibs.OpenXml
{
    internal class ExcelOpenXmlProvider : ExcelProviderBase
    {
        private IExcelReader _excelReader;
        private IExcelWriter _excelWriter;
        public ExcelOpenXmlProvider(bool printHeader)
        {
            _excelWriter = new ExcelOpenXmlSheetWriter(printHeader);
            _excelReader = new ExcelOpenXmlSheetReader();
        }


        public override IEnumerable<IDictionary<string, object>> Query(Stream stream, bool UseHeaderRow = false)
        {
             return _excelReader.Query(stream, UseHeaderRow);
        }

        public override IEnumerable<T> Query<T>(Stream stream)
        {
            return _excelReader.Query<T>(stream);
        }

        public override void SaveAs(Stream stream, object input)
        {
            _excelWriter.SaveAs(stream, input);
        }

    }
}
