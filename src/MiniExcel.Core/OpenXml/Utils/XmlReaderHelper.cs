using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;
using MiniExcelLib.Core.OpenXml.Constants;
using Zomp.SyncMethodGenerator;

namespace MiniExcelLib.Core.OpenXml.Utils;

internal static partial class XmlReaderHelper
{
    private static readonly string[] Ns = [Schemas.SpreadsheetmlXmlns, Schemas.SpreadsheetmlXmlStrictns];
    
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
                var value = await ReadStringItemAsync(reader, cancellationToken).ConfigureAwait(false);
                yield return value;
            }
            else if (!await SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
            {
                break;
            }
        }
    }

    
    /// <summary>
    /// Copied and modified from ExcelDataReader - @MIT License
    /// </summary>
    [CreateSyncVersion]
    public static async Task<string> ReadStringItemAsync(XmlReader reader, CancellationToken cancellationToken = default)
    {
        var result = new StringBuilder();
        if (!await ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
            return string.Empty;

        while (!reader.EOF)
        {
            if (IsStartElement(reader, "t", Ns))
            {
                // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                result.Append(await reader.ReadElementContentAsStringAsync()
#if NET6_0_OR_GREATER
                    .WaitAsync(cancellationToken)
#endif
                    .ConfigureAwait(false));
            }
            else if (IsStartElement(reader, "r", Ns))
            {
                result.Append(await ReadRichTextRunAsync(reader, cancellationToken).ConfigureAwait(false));
            }
            else if (!await SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
            {
                break;
            }
        }

        return result.ToString();
    }

    /// <summary>
    /// Copied and modified from ExcelDataReader - @MIT License
    /// </summary>
    [CreateSyncVersion]
    private static async Task<string> ReadRichTextRunAsync(XmlReader reader, CancellationToken cancellationToken = default)
    {
        var result = new StringBuilder();
        if (!await ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
            return string.Empty;

        while (!reader.EOF)
        {
            if (IsStartElement(reader, "t", Ns))
            {
                result.Append(await reader.ReadElementContentAsStringAsync()
#if NET6_0_OR_GREATER
                    .WaitAsync(cancellationToken)
#endif
                    .ConfigureAwait(false));
            }
            else if (!await SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
            {
                break;
            }
        }

        return result.ToString();
    }
    
    internal static XmlReaderSettings GetXmlReaderSettings(bool async) => new()
    {
        IgnoreComments = true,
        IgnoreWhitespace = true,
        XmlResolver = null,
        Async = async,
    };
}