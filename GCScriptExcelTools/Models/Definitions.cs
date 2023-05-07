using GCScriptExcelTools.Models;

namespace GCScriptExcelTools;
public class Definitions
{
    // APPLY
    public SortWorksheets? SortWorksheets { get; set; }
    public FontSettings? FontSettings { get; set; }

    // REMOVE
    public bool RemoveInvisibleWorksheets { get; set; }
    public bool RemoveEmptyWorksheets { get; set; }
    public bool RemoveFormatting { get; set; }
    public bool RemoveBackgroundColor { get; set; }
    public bool RemoveFontColor { get; set; }
    public bool GetLastRealEmptyRow { get; set; }
    public bool GetLastRealEmptyColumn { get; set; }
}