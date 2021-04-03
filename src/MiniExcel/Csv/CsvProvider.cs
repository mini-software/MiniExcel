using System.Collections.Generic;
using System.IO;

namespace MiniExcelLibs.Csv
{
    internal class CsvProvider : ExcelProviderBase
    {
        private IExcelReader _csvlReader;
        private IExcelWriter _csvWriter;
        public CsvProvider()
        {
            _csvWriter = new CsvWriter();
            _csvlReader = new CsvReader();
        }


        public override IEnumerable<IDictionary<string, object>> Query(Stream stream, bool UseHeaderRow = false)
        {
            return _csvlReader.Query(stream, UseHeaderRow);
        }

        public override IEnumerable<T> Query<T>(Stream stream)
        {
            return _csvlReader.Query<T>(stream);
        }

        public override void SaveAs(Stream stream, object input)
        {
            _csvWriter.SaveAs(stream, input);
        }

    }
}
