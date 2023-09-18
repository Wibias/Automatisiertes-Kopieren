using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using static Automatisiertes_Kopieren.FileManager.StringUtilities;

namespace Automatisiertes_Kopieren
{
    public static class ValidationHelper
    {
        public static string? DetermineProtokollbogen(double monthsAndDays)
        {
            Dictionary<double, string> protokollbogenMap = new Dictionary<double, string>
            {
        { 10.15, "Kind_Protokollbogen_12_Monate" },
        { 16.15, "Kind_Protokollbogen_18_Monate" },
        { 22.15, "Kind_Protokollbogen_24_Monate" },
        { 27.15, "Kind_Protokollbogen_30_Monate" },
        { 33.15, "Kind_Protokollbogen_36_Monate" },
        { 39.15, "Kind_Protokollbogen_42_Monate" },
        { 45.15, "Kind_Protokollbogen_48_Monate" },
        { 51.15, "Kind_Protokollbogen_54_Monate" },
        { 57.15, "Kind_Protokollbogen_60_Monate" },
        { 63.15, "Kind_Protokollbogen_66_Monate" },
        { 69.15, "Kind_Protokollbogen_72_Monate" },
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
            // Exit if no name is provided
            if (string.IsNullOrWhiteSpace(kidName))
            {
                Log.Error("Kid name is empty or whitespace.");
                MessageBox.Show("Bitte geben Sie den Namen eines Kindes an.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            string groupFolder = ConvertSpecialCharacters(groupDropdownText, ConversionType.Umlaute);  // Use the converted group name

            string groupPath = $@"{homeFolder}\Entwicklungsberichte\{groupFolder} Entwicklungsberichte\Aktuell";  // Use the passed in parameter

            // Log the constructed group path for debugging
            Log.Information($"Constructed group path: {groupPath}");

            if (!System.IO.Directory.Exists(groupPath))
            {
                Log.Error($"Group path does not exist: {groupPath}");
                MessageBox.Show($"Der Pfad für den Gruppenordner {groupFolder} ist nicht zugänglich. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            bool kidNameExists = System.IO.Directory.GetDirectories(groupPath).Any(dir => dir.Split(System.IO.Path.DirectorySeparatorChar).Last().Equals(kidName, StringComparison.OrdinalIgnoreCase));

            // Log the result of the kid name check
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
                return null;  // Handle how you'd like to exit if no year is provided.
            }

            bool isValidYear = int.TryParse(reportYearText, out int parsedYear) && parsedYear >= 2023 && parsedYear <= 2099;

            if (!isValidYear)
            {
                throw new Exception("Das Jahr muss aus genau 4 Ziffern bestehen, und zwischen 2023 und 2099 liegen. Bitte geben Sie ein gültiges Jahr ein.");
            }

            return parsedYear;
        }
    }
}