using System;
using System.IO;
using System.Threading.Tasks;
using Newtonsoft.Json;
using static Automatisiertes_Kopieren.Helper.LoggingHelper;

namespace Automatisiertes_Kopieren;

public class AppSettings
{
    private readonly string? _homeFolderPath;

    public string? HomeFolderPath
    {
        get => _homeFolderPath;
        init
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("HomeFolderPath cannot be null or whitespace.");

            _homeFolderPath = value;
        }
    }

    public static async Task SaveSettingsAsync(AppSettings settings)
    {
        try
        {
            var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Automatisiertes_Kopieren", "settings.json");

            LogMessage($"Attempting to save settings to: {path}", LogLevel.Warning);

            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            await File.WriteAllTextAsync(path, json);

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