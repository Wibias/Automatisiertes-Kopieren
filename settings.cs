using Newtonsoft.Json;
using System;
using System.IO;
using System.Threading.Tasks;
using static Automatisiertes_Kopieren.Helper.LoggingHelper;

namespace Automatisiertes_Kopieren
{
    public class Settings
    {
        public string? HomeFolderPath { get; init; }

        public static async Task SaveSettingsAsync(Settings settings)
        {
            try
            {
                var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                var directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Automatisiertes_Kopieren");
                var path = Path.Combine(directoryPath, "settings.json");

                LogMessage($"Versuche Einstellung hier zu speichern: {path }", LogLevel.Warning);

                Directory.CreateDirectory(directoryPath );
                await File.WriteAllTextAsync(directoryPath, json);

                LogMessage("Settings saved successfully.");
            }
            catch (Exception ex)
            {
                LogMessage($"Fehler beim Speichern von Einstellungen: {ex.Message}", LogLevel.Warning);
            }
        }

        public static Settings? LoadSettings()
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Automatisiertes_Kopieren", "settings.json");

            if (!File.Exists(path)) return new Settings();

            try
            {
                var json = File.ReadAllText(path);
                return JsonConvert.DeserializeObject<Settings>(json);

            }
            catch (Exception ex)
            {
                LogMessage($"Error while loading settings: {ex.Message}", LogLevel.Warning);
                return null;
            }
        }
    }
}