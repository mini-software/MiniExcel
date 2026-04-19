using MiniExcelLib.OpenXml.Constants;

namespace MiniExcelLib.OpenXml.Utils;

internal static partial class XmlReaderHelper
{
    private static readonly string[] Ns = [Schemas.SpreadsheetmlXmlMain, Schemas.SpreadsheetmlXmlStrictNs];

    // Copied and modified from ExcelDataReader - @MIT License
    [CreateSyncVersion]
    public static async Task<bool> ReadFirstContentAsync(this XmlReader reader, CancellationToken cancellationToken = default)
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

    // Copied and modified from ExcelDataReader - @MIT License
    [CreateSyncVersion]
    public static async Task<bool> SkipContentAsync(this XmlReader reader, CancellationToken cancellationToken = default)
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

    /// <summary>
    /// e.g skip row 1 to row 2
    /// </summary>
    [CreateSyncVersion]
    public static async Task SkipToNextSiblingAsync(this XmlReader reader, CancellationToken cancellationToken = default)
    {
        while (!reader.EOF)
        {
            if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                break;
        }
    }

    public static bool IsStartElement(this XmlReader reader, string name, params string[] nss)
    {
        return nss.Any(s => reader.IsStartElement(name, s));
    }

    public static string? GetAttribute(this XmlReader reader, string name, params string[] nss)
    {
        return nss
            .Select(ns => reader.GetAttribute(name, ns))
            .FirstOrDefault(at => at is not null);
    }

    // Copied and modified from ExcelDataReader - @MIT License
    [CreateSyncVersion]
    public static async Task<string> ReadStringItemAsync(this XmlReader reader, CancellationToken cancellationToken = default)
    {
        var result = new StringBuilder();
        if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
            return string.Empty;

        while (!reader.EOF)
        {
            if (reader.IsStartElement("t", Ns))
            {
                // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                result.Append(await reader.ReadElementContentAsStringAsync()
#if NET6_0_OR_GREATER
                    .WaitAsync(cancellationToken)
#endif
                    .ConfigureAwait(false));
            }
            else if (reader.IsStartElement("r", Ns))
            {
                result.Append(await reader.ReadRichTextRunAsync(cancellationToken).ConfigureAwait(false));
            }
            else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
            {
                break;
            }
        }

        return result.ToString();
    }

    // Copied and modified from ExcelDataReader - @MIT License
    [CreateSyncVersion]
    private static async Task<string> ReadRichTextRunAsync(this XmlReader reader, CancellationToken cancellationToken = default)
    {
        var result = new StringBuilder();
        if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
            return string.Empty;

        while (!reader.EOF)
        {
            if (reader.IsStartElement("t", Ns))
            {
                result.Append(await reader.ReadElementContentAsStringAsync()
#if NET6_0_OR_GREATER
                        .WaitAsync(cancellationToken)
#endif
                    .ConfigureAwait(false));
            }
            else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
            {
                break;
            }
        }

        return result.ToString();
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
        if (!reader.IsStartElement("sst", nss))
            yield break;

        if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
            yield break;

        while (!reader.EOF)
        {
            if (reader.IsStartElement("si", nss))
            {
                var value = await reader.ReadStringItemAsync(cancellationToken).ConfigureAwait(false);
                yield return value;
            }
            else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
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
        Async = async
    };
}
