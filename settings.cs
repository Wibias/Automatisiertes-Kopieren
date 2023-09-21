using Serilog;
using System;

namespace Automatisiertes_Kopieren
{
    public class AppSettings
    {
        public string? HomeFolderPath { get; set; }

        public void SaveSettings(AppSettings settings)
        {
            try
            {
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(settings, Newtonsoft.Json.Formatting.Indented);
                string path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Automatisiertes_Kopieren", "settings.json");

                Log.Information($"Attempting to save settings to: {path}");

                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path)!);
                System.IO.File.WriteAllText(path, json);

                Log.Information("Settings saved successfully.");
            }
            catch (Exception ex)
            {
                Log.Error($"Error saving settings: {ex.Message}");
            }
        }

        public AppSettings? LoadSettings()
        {
            string path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Automatisiertes_Kopieren", "settings.json");
            if (System.IO.File.Exists(path))
            {
                string json = System.IO.File.ReadAllText(path);
                return Newtonsoft.Json.JsonConvert.DeserializeObject<AppSettings>(json);
            }
            return new AppSettings();
        }

    }
}
