using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using Newtonsoft.Json;
using static Automatisiertes_Kopieren.Helper.LoggingHelper;

namespace Automatisiertes_Kopieren;

public class Settings
{
    private readonly string? _homeFolderPath;

    private string PreferredLanguage { get; set; } = CultureInfo.InstalledUICulture.Name;

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

    public static async Task SaveSettingsAsync(Settings settings)
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

    public static Settings? LoadSettings()
    {
        var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Automatisiertes_Kopieren", "settings.json");
        if (!File.Exists(path)) return new Settings();
        var json = File.ReadAllText(path);
        return JsonConvert.DeserializeObject<Settings>(json);
    }
}