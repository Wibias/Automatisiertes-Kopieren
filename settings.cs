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
                throw new ArgumentException("Hauptverzeichnis kann nicht null oder ein Leerzeichen sein.");

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

            LogMessage($"Versuche Einstellung hier zu speichern: {path}", LogLevel.Warning);

            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            await File.WriteAllTextAsync(path, json);

            LogMessage("Settings saved successfully.");
        }
        catch (Exception ex)
        {
            LogMessage($"Fehler beim Speichern von Einstellungen: {ex.Message}", LogLevel.Warning);
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