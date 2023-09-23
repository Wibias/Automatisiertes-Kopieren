using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using static Automatisiertes_Kopieren.FileManager.StringUtilities;

namespace Automatisiertes_Kopieren
{
    public static class ValidationHelper
    {
        public static bool IsValidPath(string path)
        {
            if (!Directory.Exists(path))
                return false;

            try
            {
                Path.GetFullPath(path);
            }
            catch
            {
                return false;
            }

            string tempFile = Path.Combine(path, "tempFileToCheckWritePermission.txt");
            try
            {
                using (FileStream fs = File.Create(tempFile, 1, FileOptions.DeleteOnClose))
                {
                    // Do nothing, just create the file and close it
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        public static (string directoryPath, string fileName)? DetermineProtokollbogen(double monthsAndDays)
        {
            Dictionary<double, (string directoryPath, string fileName)> protokollbogenMap = new Dictionary<double, (string, string)>
            {
                { 10.15, (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_12_Monate") },
                { 16.15, (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_16_Monate") },
                { 22.15, (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_24_Monate") },
                { 27.15, (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_30_Monate") },
                { 33.15, (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_36_Monate") },
                { 39.15, (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_42_Monate") },
                { 45.15, (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_48_Monate") },
                { 51.15, (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_54_Monate") },
                { 57.15, (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_60_Monate") },
                { 63.15, (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_66_Monate") },
                { 69.15, (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_72_Monate") },
            };

            foreach (var entry in protokollbogenMap.OrderByDescending(kvp => kvp.Key))
            {
                if (monthsAndDays >= entry.Key)
                {
                    return entry.Value;
                }
            }

            Serilog.Log.Warning($"Kein Protokollbogen für folgenden Monatswert gefunden: {monthsAndDays}");
            return null;
        }

        public static string? ValidateKidName(string kidName, string homeFolder, string groupDropdownText)
        {
            if (string.IsNullOrWhiteSpace(kidName))
            {
                Log.Error("Kid name is empty or whitespace.");
                MessageBox.Show("Bitte geben Sie den Namen eines Kindes an.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            string groupFolder = ConvertSpecialCharacters(groupDropdownText, ConversionType.Umlaute);

            string groupPath = $@"{homeFolder}\Entwicklungsberichte\{groupFolder} Entwicklungsberichte\Aktuell";

            if (!IsValidPath(groupPath))
            {
                Log.Error($"Der Gruppenpfad ist nicht gültig oder zugänglich: {groupPath}");
                MessageBox.Show($"Der Pfad für den Gruppenordner {groupFolder} ist nicht zugänglich. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            bool kidNameExists = System.IO.Directory.GetDirectories(groupPath).Any(dir => dir.Split(System.IO.Path.DirectorySeparatorChar).Last().Equals(kidName, StringComparison.OrdinalIgnoreCase));

            Log.Information($"Kid name exists in group directory: {kidNameExists}");

            if (!kidNameExists)
            {
                Log.Error($"Kid name not found in group directory: {kidName}");
                MessageBox.Show($"Der Name des Kindes wurde im Gruppenverzeichnis nicht gefunden. Bitte geben Sie einen gültigen Namen an.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            return kidName;
        }

        public static int? ValidateReportYearFromTextbox(string reportYearText)
        {
            if (string.IsNullOrWhiteSpace(reportYearText))
            {
                return null;
            }

            if (!int.TryParse(reportYearText, out int parsedYear) || parsedYear < 2023 || parsedYear > 2099)
            {
                throw new ArgumentException("Das Jahr muss aus genau 4 Ziffern bestehen, und zwischen 2023 und 2099 liegen. Bitte geben Sie ein gültiges Jahr ein.");
            }

            return parsedYear;
        }
    }
}