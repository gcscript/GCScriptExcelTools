using System.Text.Json;

namespace GCScriptExcelTools.Models;
public class Preset
{
    public string Title { get; set; } = string.Empty;
    public PresetRemoveOptions Remove { get; set; } = new ();
    public PresetOthersOptions Others { get; set; } = new();
}

public class PresetRemoveOptions
{
    public bool HiddenWorksheets { get; set; }
    public bool EmptyWorksheets { get; set; }
    public bool EmptyRows { get; set; }
    public bool EmptyColumns { get; set; }
    public bool HiddenRows { get; set; }
    public bool FontColor { get; set; }
    public bool BackgroundColor { get; set; }
    public bool Formatting { get; set; }
    public List<RemoveColumns>? Columns { get; set; }
}

public class PresetOthersOptions
{
    public bool GetLastRealEmptyRow { get; set; }
    public bool GetLastRealEmptyColumn { get; set; }
    public FindHeader? FindHeader { get; set; }
}