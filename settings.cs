using System;
using System.IO;
using Newtonsoft.Json;
using static Automatisiertes_Kopieren.LoggingService;

namespace Automatisiertes_Kopieren;

public class AppSettings
{
    private static readonly LoggingService LoggingService = new();
    public string? HomeFolderPath { get; init; }

    public static void SaveSettings(AppSettings settings)
    {
        try
        {
            var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Automatisiertes_Kopieren", "settings.json");

            LogMessage($"Attempting to save settings to: {path}", LogLevel.Warning);

            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, json);

            LogMessage("Settings saved successfully.");
        }
        catch (Exception ex)
        {
            LogMessage($"Error saving settings: {ex.Message}", LogLevel.Warning);
        }
    }

    public static AppSettings? LoadSettings()
    {
        var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Automatisiertes_Kopieren", "settings.json");
        if (!File.Exists(path)) return new AppSettings();
        var json = File.ReadAllText(path);
        return JsonConvert.DeserializeObject<AppSettings>(json);
    }
}