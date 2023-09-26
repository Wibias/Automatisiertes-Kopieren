using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using static Automatisiertes_Kopieren.FileManager.StringUtilities;

namespace Automatisiertes_Kopieren
{
    public static class ValidationHelper
    {
        private readonly static LoggingService _loggingService = new LoggingService();
        public static (string directoryPath, string fileName)? DetermineProtokollbogen(double monthsAndDays)
        {
            _loggingService.LogMessage($"Checking value against ranges: {monthsAndDays}", LoggingService.LogLevel.Information);
            var protokollbogenRanges = new List<(double start, double end, (string directoryPath, string fileName) value)>
            {
                (10.15, 16.14, (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_12_Monate")),
                (16.15, 22.14, (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_16_Monate")),
                (22.15, 27.14, (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_24_Monate")),
                (27.15, 33.14, (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_30_Monate")),
                (33.15, 39.14, (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_36_Monate")),
                (39.15, 45.14, (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_42_Monate")),
                (45.15, 51.14, (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_48_Monate")),
                (51.15, 57.14, (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_54_Monate")),
                (57.15, 63.14, (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_60_Monate")),
                (63.15, 69.14, (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_66_Monate")),
                (69.15, 84.00, (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_72_Monate")),
            };


            foreach (var range in protokollbogenRanges.OrderByDescending(r => r.start))
            {
                if (monthsAndDays >= range.start && monthsAndDays <= range.end)
                {
                    return range.value;
                }
            }

            _loggingService.LogAndShowMessage($"No Protokollbogen for Month value found: {monthsAndDays}",
                                              $"Kein Protokollbogen für folgenden Monatswert gefunden: {monthsAndDays}",
                                              LoggingService.LogLevel.Warning);
            MainWindow.OperationState.OperationsSuccessful = false;
            return null;
        }

        public static double ConvertToDecimalFormat(double monthsAndDays)
        {
            string monthsAndDaysRaw = monthsAndDays.ToString("0.00", CultureInfo.InvariantCulture);

            if (double.TryParse(monthsAndDaysRaw, out double parsedValue))
            {
                return parsedValue;
            }
            else
            {
                return 0;
            }
        }

        public static string? ValidateKidName(string kidName, string homeFolder, string groupDropdownText)
        {
            if (string.IsNullOrWhiteSpace(kidName))
            {
                _loggingService.LogAndShowMessage("Kid name is empty or whitespace.",
                                                  "Bitte geben Sie den Namen eines Kindes an.");
                return null;
            }

            string groupFolder = ConvertSpecialCharacters(groupDropdownText, ConversionType.Umlaute);

            string groupPath = $@"{homeFolder}\Entwicklungsberichte\{groupFolder} Entwicklungsberichte\Aktuell";


            if (!Directory.Exists(groupPath))
            {
                _loggingService.LogAndShowMessage($"Group path does not exist: {groupPath}",
                                                  $"Der Pfad für den Gruppenordner {groupFolder} ist nicht zugänglich. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut.");
                return null;
            }

            bool kidNameExists = Directory.GetDirectories(groupPath).Any(dir => dir.Split(System.IO.Path.DirectorySeparatorChar).Last().Equals(kidName, StringComparison.OrdinalIgnoreCase));

            if (!kidNameExists)
            {
                _loggingService.LogAndShowMessage($"Kid name not found in group directory: {kidName}",
                                                  $"Der Name des Kindes wurde im Gruppenverzeichnis nicht gefunden. Bitte geben Sie einen gültigen Namen an.");
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

            bool isValidYear = int.TryParse(reportYearText, out int parsedYear) && parsedYear >= 2023 && parsedYear <= 2099;

            if (!isValidYear)
            {
                throw new Exception("Das Jahr muss aus genau 4 Ziffern bestehen, und zwischen 2023 und 2099 liegen. Bitte geben Sie ein gültiges Jahr ein.");
            }

            return parsedYear;
        }
    }
}