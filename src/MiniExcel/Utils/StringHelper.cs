using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using MiniExcelLibs.OpenXml.Constants;
using Zomp.SyncMethodGenerator;

namespace MiniExcelLibs.Utils;

internal static partial class StringHelper
{
    private static readonly string[] Ns = [Schemas.SpreadsheetmlXmlns, Schemas.SpreadsheetmlXmlStrictns];

    public static string GetLetters(string content) => new([..content.Where(char.IsLetter)]);
    public static int GetNumber(string content) => int.Parse(new string([..content.Where(char.IsNumber)]));

    /// <summary>
    /// Copied and modified from ExcelDataReader - @MIT License
    /// </summary>
    [CreateSyncVersion]
    public static async Task<string> ReadStringItemAsync(XmlReader reader, CancellationToken cancellationToken = default)
    {
        var result = new StringBuilder();
        if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
            return string.Empty;

        while (!reader.EOF)
        {
            if (XmlReaderHelper.IsStartElement(reader, "t", Ns))
            {
                // There are multiple <t> in a <si>. Concatenate <t> within an <si>.
                result.Append(await reader.ReadElementContentAsStringAsync()
#if NET6_0_OR_GREATER
                    .WaitAsync(cancellationToken)
#endif
                    .ConfigureAwait(false));
            }
            else if (XmlReaderHelper.IsStartElement(reader, "r", Ns))
            {
                result.Append(await ReadRichTextRunAsync(reader, cancellationToken).ConfigureAwait(false));
            }
            else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
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
        if (!await XmlReaderHelper.ReadFirstContentAsync(reader, cancellationToken).ConfigureAwait(false))
            return string.Empty;

        while (!reader.EOF)
        {
            if (XmlReaderHelper.IsStartElement(reader, "t", Ns))
            {
                result.Append(await reader.ReadElementContentAsStringAsync()
#if NET6_0_OR_GREATER
                    .WaitAsync(cancellationToken)
#endif
                    .ConfigureAwait(false));
            }
            else if (!await XmlReaderHelper.SkipContentAsync(reader, cancellationToken).ConfigureAwait(false))
            {
                break;
            }
        }

        return result.ToString();
    }
}