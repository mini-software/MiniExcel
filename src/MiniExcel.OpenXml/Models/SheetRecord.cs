namespace MiniExcelLib.OpenXml.Models;

internal sealed class SheetRecord(string name, string state, uint id, string rid, bool active)
{
    public string Name { get; } = name;
    public string State { get; set; } = state;
    public uint Id { get; } = id;
    public string Rid { get; set; } = rid;
    public string Path { get; set; }
    public bool Active { get; } = active;

    public SheetInfo ToSheetInfo(uint index)
    {
        if (string.IsNullOrEmpty(State))
            return new SheetInfo(Id, index, Name, SheetState.Visible, Active);
        
        if (Enum.TryParse(State, true, out SheetState stateEnum))
            return new SheetInfo(Id, index, Name, stateEnum, Active);
        
        throw new ArgumentException($"Unable to parse sheet state. Sheet name: {Name}");
    }
}