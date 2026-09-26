namespace MiniExcelLib.OpenXml.Models;

internal class FileDto
{
    internal int SheetIndex { get; set; }
    internal int RowIndex { get; set; }
    internal int CellIndex { get; set; }

    /// <summary>
    /// Disambiguates the generated media and relationship ids when multiple image values share the same
    /// anchor cell. Left unset by the regular <c>SaveAs</c> pipeline, which never places two images on
    /// one cell.
    /// </summary>
    internal string? IdSuffix { get; set; }

    internal string Id => string.IsNullOrEmpty(IdSuffix)
        ? $"rFileId_{SheetIndex}_{RowIndex + 1}_{CellIndex + 1}"
        : $"rFileId_{SheetIndex}_{RowIndex + 1}_{CellIndex + 1}_{IdSuffix}";
    internal string Path => $"xl/media/{Id}.{Extension}";
    internal bool IsImage { get; set; }
    internal string Extension { get; set; }
    internal byte[] Contents { get; set; }

    /// <summary>
    /// Anchor size in EMUs. When unset, the drawing falls back to the default image size.
    /// </summary>
    internal long? ImageWidthEmu { get; set; }
    internal long? ImageHeightEmu { get; set; }
}