using GCScriptExcelTools.Models;
using System.Text.Json;

namespace GCScriptExcelTools.Services;

public static class Services
{
    public static void SavePreset(Preset preset)
    {
        string path = Path.Combine(Settings.PresetsPath, $"{preset.Title}.json");

        if (!Directory.Exists(Path.GetDirectoryName(path)))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
        }

        if (File.Exists(path))
        {
            File.Delete(path);
        }

        var json = JsonSerializer.Serialize(preset, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(path, json);
    }
}
