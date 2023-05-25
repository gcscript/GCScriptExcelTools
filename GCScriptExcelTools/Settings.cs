namespace GCScriptExcelTools;
public static class Settings
{
    public static string AppPath { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
    public static string PresetsPath { get; set; } = Path.Combine(AppPath, "Presets");
}
