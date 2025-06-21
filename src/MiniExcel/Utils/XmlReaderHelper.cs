using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Zomp.SyncMethodGenerator;

namespace MiniExcelLibs.Utils;

internal static partial class XmlReaderHelper
{
    /// <summary>
    /// Pass &lt;?xml&gt; and &lt;worksheet&gt;
    /// </summary>
    [CreateSyncVersion]
    public static async Task PassXmlDeclarationAndWorksheetAsync(this XmlReader reader, CancellationToken cancellationToken = default)
    {
        await reader.MoveToContentAsync()
#if NET6_0_OR_GREATER
            .WaitAsync(cancellationToken)
#endif
            .ConfigureAwait(false);
        await reader.ReadAsync()
#if NET6_0_OR_GREATER
            .WaitAsync(cancellationToken)
#endif
            .ConfigureAwait(false);
    }

    /// <summary>
    /// e.g skip row 1 to row 2
    /// </summary>
    [CreateSyncVersion]
    public static async Task SkipToNextSameLevelDomAsync(XmlReader reader, CancellationToken cancellationToken = default)
    {
        while (!reader.EOF)
        {
            if (!await SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
                break;
        }
    }

    //Method from ExcelDataReader @MIT License
    [CreateSyncVersion]
    public static async Task<bool> ReadFirstContentAsync(XmlReader reader, CancellationToken cancellationToken = default)
    {
        if (reader.IsEmptyElement)
        {
            await reader.ReadAsync()
#if NET6_0_OR_GREATER
                .WaitAsync(cancellationToken)
#endif
                .ConfigureAwait(false);
            return false;
        }

        await reader.MoveToContentAsync()
#if NET6_0_OR_GREATER
            .WaitAsync(cancellationToken)
#endif
            .ConfigureAwait(false);
        await reader.ReadAsync()
#if NET6_0_OR_GREATER
            .WaitAsync(cancellationToken)
#endif
            .ConfigureAwait(false);
        return true;
    }

    //Method from ExcelDataReader @MIT License
    [CreateSyncVersion]
    public static async Task<bool> SkipContentAsync(XmlReader reader, CancellationToken cancellationToken = default)
    {
        if (reader.NodeType == XmlNodeType.EndElement)
        {
            await reader.ReadAsync()
#if NET6_0_OR_GREATER
                .WaitAsync(cancellationToken)
#endif
                .ConfigureAwait(false);
            return false;
        }

        await reader.SkipAsync()
#if NET6_0_OR_GREATER
            .WaitAsync(cancellationToken)
#endif
            .ConfigureAwait(false);
        return true;
    }

    public static bool IsStartElement(XmlReader reader, string name, params string[] nss)
    {
        return nss.Any(s => reader.IsStartElement(name, s));
    }

    public static string? GetAttribute(XmlReader reader, string name, params string[] nss)
    {
        return nss
            .Select(ns => reader.GetAttribute(name, ns))
            .FirstOrDefault(at => at is not null);
    }

    [CreateSyncVersion]
    public static async IAsyncEnumerable<string> GetSharedStringsAsync(Stream stream, [EnumeratorCancellation]CancellationToken cancellationToken = default, params string[] nss)
    {
        var xmlSettings = GetXmlReaderSettings(
#if SYNC_ONLY
            false
#else
            true
#endif
        );
        
        using var reader = XmlReader.Create(stream, xmlSettings);
        if (!IsStartElement(reader, "sst", nss))
            yield break;

        if (!await ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
            yield break;

        while (!reader.EOF)
        {
            if (IsStartElement(reader, "si", nss))
            {
                var value = await StringHelper.ReadStringItemAsync(reader, cancellationToken).ConfigureAwait(false);
                yield return value;
            }
            else if (!await SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
            {
                break;
            }
        }
    }

    internal static XmlReaderSettings GetXmlReaderSettings(bool async) => new()
    {
        IgnoreComments = true,
        IgnoreWhitespace = true,
        XmlResolver = null,
        Async = async,
    };
}