namespace GCScriptExcelTools;
public static class Settings
{
    public static string AppPath { get; } = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)!;
    public static string PresetsPath { get; set; } = Path.Combine(AppPath, "Presets");
}
