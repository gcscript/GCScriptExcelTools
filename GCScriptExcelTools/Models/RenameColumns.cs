namespace GCScriptExcelTools.Models;

public enum ERenameColumnsFilterType { Equals, Contains, StartsWith, EndsWith }
public class RenameColumns
{
    public ERenameColumnsFilterType Filter { get; set; }
    public string Find { get; set; }
    public string Replace { get; set; }
}
