using System.Text.Json;
using GCScriptExcelTools.Models;

namespace GCScriptExcelTools.Services;

public static class FPreset
{
    public static bool Save(Preset preset)
    {
        try
        {
            string path = Path.Combine(Settings.PresetsPath, $"{preset.Title}.json");
            if (!Directory.Exists(Path.GetDirectoryName(path))) { Directory.CreateDirectory(Path.GetDirectoryName(path)!); }
            if (File.Exists(path)) { File.Delete(path); }
            var json = JsonSerializer.Serialize(preset, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
            return true;
        }
        catch { return false; }
    }

    public static Preset? Load(string path)
    {
        try
        {
            if (!File.Exists(path)) { return null; }
            var json = File.ReadAllText(path);
            if (string.IsNullOrEmpty(json)) { return null; }
            return JsonSerializer.Deserialize<Preset>(json);
        }
        catch { return null; }
    }

    public static bool Remove(string path)
    {
        try
        {
            if (!File.Exists(path)) { return true; }
            File.Delete(path);
            return true;
        }
        catch { return false; }
    }
}
