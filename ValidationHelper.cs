using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

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

            Serilog.Log.Warning($"No Protokollbogen found for month value: {monthsAndDays}");
            return null;
        }

        public static string? ValidateKidName(string kidName, string homeFolder, string groupDropdownText)
        {
            // Exit if no name is provided
            if (string.IsNullOrWhiteSpace(kidName))
            {
                MessageBox.Show("Please provide a kid's name.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;  // You can choose how to handle this case.
            }

            string groupFolder = groupDropdownText;  // Use the passed in parameter

            string groupPath = $@"{homeFolder}\Entwicklungsberichte\{groupFolder}\Aktuell";  // Use the passed in parameter
            if (!System.IO.Directory.Exists(groupPath))
            {
                MessageBox.Show($"The path for the group folder {groupFolder} is not accessible. Please check the path and try again.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            bool kidNameExists = System.IO.Directory.GetDirectories(groupPath).Any(dir => dir.Split(System.IO.Path.DirectorySeparatorChar).Last().Equals(kidName, StringComparison.OrdinalIgnoreCase));

            if (!kidNameExists)
            {
                MessageBox.Show($"The kid's name was not found in the group directory. Please provide a valid name.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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