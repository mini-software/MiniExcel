namespace MiniExcelLib.OpenXml.Models;

public class SheetInfo(uint id, uint index, string name, SheetState sheetState, bool active)
{
    /// <summary>
    /// Internal sheet id - depends on the order in which the sheet is added.
    /// </summary>
    public uint Id { get; } = id;

    /// <summary>
    /// The 0-based index of the worksheet in the workbook
    /// </summary>
    public uint Index { get; } = index;

    /// <summary>
    /// The name of the worksheet
    /// </summary>
    public string Name { get; } = name;

    /// <summary>
    /// Sheet visibility state
    /// </summary>
    public SheetState State { get; } = sheetState;

    /// <summary>
    /// Indicates whether the worksheet was active the last time 
    /// </summary>
    public bool Active { get; } = active;
}

public enum SheetState { Visible, Hidden, VeryHidden }