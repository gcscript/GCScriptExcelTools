namespace GCScriptExcelTools.Models;

public enum ERemoveColumnsFilterType { Equals, Contains, StartsWith, EndsWith }
public class RemoveColumns
{
    public ERemoveColumnsFilterType Filter { get; set; }
    public string Keyword { get; set; }
}
