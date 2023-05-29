using GCScriptExcelTools.Models;

namespace GCScriptExcelTools;
public class Definitions
{
    public SortWorksheets? SortWorksheets { get; set; }
    public FontSettings? FontSettings { get; set; }
    public FindHeader? FindHeader { get; set; }
    public List<RemoveColumns>? RemoveColumnsList { get; set; }
    public List<RenameColumns>? RenameColumnsList { get; set; }
    public bool RemoveHiddenWorksheets { get; set; }
    public bool RemoveHiddenRows { get; set; }
    public bool RemoveEmptyWorksheets { get; set; }
    public bool RemoveFormatting { get; set; }
    public bool RemoveBackgroundColor { get; set; }
    public bool RemoveFontColor { get; set; }
    public bool GetLastRealEmptyRow { get; set; }
    public bool GetLastRealEmptyColumn { get; set; }
}