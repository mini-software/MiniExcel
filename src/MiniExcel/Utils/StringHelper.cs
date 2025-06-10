using MiniExcelLibs.OpenXml;
using System;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace MiniExcelLibs.Utils
{
    internal static partial class StringHelper
    {
        private static readonly string[] _ns = { Config.SpreadsheetmlXmlns, Config.SpreadsheetmlXmlStrictns };

        public static string GetLetters(string content) => new string(content.Where(char.IsLetter).ToArray());
        public static int GetNumber(string content) => int.Parse(new string(content.Where(char.IsNumber).ToArray()));

        /// <summary>
        /// Copied and modified from ExcelDataReader - @MIT License
        /// </summary>
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<string> ReadStringItemAsync(XmlReader reader, CancellationToken ct = default)
        {
            var result = new StringBuilder();
            if (!await XmlReaderHelper.ReadFirstContentAsync(reader, ct))
                return string.Empty;

            while (!reader.EOF)
            {
                if (XmlReaderHelper.IsStartElement(reader, "t", _ns))
                {
                    // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                    result.Append(await reader.ReadElementContentAsStringAsync()
#if NET6_0_OR_GREATER
                        .WaitAsync(ct)
#endif
                        );
                }
                else if (XmlReaderHelper.IsStartElement(reader, "r", _ns))
                {
                    result.Append(await ReadRichTextRunAsync(reader, ct));
                }
                else if (!await XmlReaderHelper.SkipContentAsync(reader, ct))
                {
                    break;
                }
            }

            return result.ToString();
        }

        /// <summary>
        /// Copied and modified from ExcelDataReader - @MIT License
        /// </summary>
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private static async Task<string> ReadRichTextRunAsync(XmlReader reader, CancellationToken ct = default)
        {
            var result = new StringBuilder();
            if (!await XmlReaderHelper.ReadFirstContentAsync(reader, ct))
                return string.Empty;

            while (!reader.EOF)
            {
                if (XmlReaderHelper.IsStartElement(reader, "t", _ns))
                {
                    result.Append(await reader.ReadElementContentAsStringAsync()
#if NET6_0_OR_GREATER
                        .WaitAsync(ct)
#endif
);
                }
                else if (!await XmlReaderHelper.SkipContentAsync(reader, ct))
                {
                    break;
                }
            }

            return result.ToString();
        }
    }
}
