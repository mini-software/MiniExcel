using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace MiniExcelLibs.Utils
{
    internal static partial class XmlReaderHelper
    {
        /// <summary>
        /// Pass &lt;?xml&gt; and &lt;worksheet&gt;
        /// </summary>
        /// <param name="reader"></param>
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task PassXmlDeclarationAndWorksheet(this XmlReader reader, CancellationToken ct = default)
        {
            await reader.MoveToContentAsync()
#if NET6_0_OR_GREATER
                        .WaitAsync(ct)
#endif
;
            await reader.ReadAsync()
#if NET6_0_OR_GREATER
                        .WaitAsync(ct)
#endif
;
        }

        /// <summary>
        /// e.g skip row 1 to row 2
        /// </summary>
        /// <param name="reader"></param>
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task SkipToNextSameLevelDomAsync(XmlReader reader, CancellationToken ct = default)
        {
            while (!reader.EOF)
            {
                if (!await SkipContentAsync(reader, ct))
                    break;
            }
        }

        //Method from ExcelDataReader @MIT License
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<bool> ReadFirstContentAsync(XmlReader reader, CancellationToken ct = default)
        {
            if (reader.IsEmptyElement)
            {
                await reader.ReadAsync()
#if NET6_0_OR_GREATER
                        .WaitAsync(ct)
#endif
;
                return false;
            }

            await reader.MoveToContentAsync()
#if NET6_0_OR_GREATER
                        .WaitAsync(ct)
#endif
                ;
            await reader.ReadAsync()
#if NET6_0_OR_GREATER
                        .WaitAsync(ct)
#endif
                ;
            return true;
        }

        //Method from ExcelDataReader @MIT License
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<bool> SkipContentAsync(XmlReader reader, CancellationToken ct = default)
        {
            if (reader.NodeType == XmlNodeType.EndElement)
            {
                await reader.ReadAsync()
#if NET6_0_OR_GREATER
                        .WaitAsync(ct)
#endif
;
                return false;
            }

            await reader.SkipAsync()
#if NET6_0_OR_GREATER
                        .WaitAsync(ct)
#endif
 ;
            return true;
        }

        public static bool IsStartElement(XmlReader reader, string name, params string[] nss)
        {
            return nss.Any(s => reader.IsStartElement(name, s));
        }

        public static string GetAttribute(XmlReader reader, string name, params string[] nss)
        {
            return nss
                .Select(ns => reader.GetAttribute(name, ns))
                .FirstOrDefault(at => at != null);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async IAsyncEnumerable<string> GetSharedStringsAsync(Stream stream, CancellationToken ct = default, params string[] nss)
        {
            using (var reader = XmlReader.Create(stream))
            {
                if (!IsStartElement(reader, "sst", nss))
                    yield break;

                if (!await ReadFirstContentAsync(reader, ct))
                    yield break;

                while (!reader.EOF)
                {
                    if (IsStartElement(reader, "si", nss))
                    {
                        var value = StringHelper.ReadStringItem(reader);
                        yield return value;
                    }
                    else if (!await SkipContentAsync(reader, ct))
                    {
                        break;
                    }
                }
            }
        }
    }
}
