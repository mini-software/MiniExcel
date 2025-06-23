namespace MiniExcelLibs.OpenXml;

public class SheetInfo(uint id, uint index, string name, SheetState sheetState, bool active)
{
    /// <summary>
    /// Internal sheet id - depends on the order in which the sheet is added
    /// </summary>
    public uint Id { get; } = id;

    /// <summary>
    /// Next sheet index - numbered from 0
    /// </summary>
    public uint Index { get; } = index;

    /// <summary>
    /// Sheet name
    /// </summary>
    public string Name { get; } = name;

    /// <summary>
    /// Sheet visibility state
    /// </summary>
    public SheetState State { get; } = sheetState;

    /// <summary>
    /// Indicates whether the sheet is active
    /// </summary>
    public bool Active { get; } = active;
}

public enum SheetState { Visible, Hidden, VeryHidden }