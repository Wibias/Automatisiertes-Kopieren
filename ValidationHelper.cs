using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static Automatisiertes_Kopieren.FileManager.StringUtilities;

namespace Automatisiertes_Kopieren
{
    public class ValidationHelper
    {
        private readonly MainWindow _mainWindow;
        private readonly LoggingService _loggingService;

        public ValidationHelper(MainWindow mainWindow)
        {
            _mainWindow = mainWindow ?? throw new ArgumentNullException(nameof(mainWindow));
            _loggingService = new LoggingService(mainWindow);
        }
        public static bool IsValidDirectoryPath(string path)
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

        public static bool IsValidFilePath(string? path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                return false;

            try
            {
                Path.GetFullPath(path);
            }
            catch
            {
                return false;
            }

            return true;
        }

        public (string directoryPath, string fileName)? DetermineProtokollbogen(double monthsAndDays)
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

            _loggingService.LogAndShowError($"Kein Protokollbogen für folgenden Monatswert gefunden: {monthsAndDays}", "Ein Fehler ist aufgetreten. Bitte überprüfen Sie die Eingaben.");
            return null;
        }

        public bool IsValidInput()
        {
            return IsHomeFolderSet() && AreAllRequiredFieldsFilled() && IsKidNameValid() != null && IsReportYearValid() != null;
        }

        public bool IsHomeFolderSet()
        {
            if (string.IsNullOrEmpty(_mainWindow.HomeFolder))
            {
                _loggingService.LogAndShowError("HomeFolder is not set.", "Bitte wählen Sie zunächst das Hauptverzeichnis aus.");
                return false;
            }
            return true;
        }

        public string? IsKidNameValid()
        {
            if (_mainWindow.HomeFolder == null)
            {
                _loggingService.LogAndShowError("HomeFolder is null during kid name validation.", "Bitte wählen Sie zunächst das Hauptverzeichnis aus.");
                return null;
            }
            string kidName = _mainWindow.kidNameComboBox.Text;

            if (string.IsNullOrWhiteSpace(kidName))
            {
                _loggingService.LogAndShowError("Der Kinder-Name ist leer oder enthält ein Leerzeichen.", "Bitte geben Sie den Namen eines Kindes an.");
                return null;
            }

            string groupFolder = ConvertSpecialCharacters(_mainWindow.groupDropdown.Text, ConversionType.Umlaute);
            string groupPath = $@"{_mainWindow.HomeFolder}\Entwicklungsberichte\{groupFolder} Entwicklungsberichte\Aktuell";

            if (!IsValidDirectoryPath(groupPath))
            {
                _loggingService.LogAndShowError($"Der Gruppenpfad ist nicht gültig oder zugänglich: {groupPath}", $"Der Pfad für den Gruppenordner {groupFolder} ist nicht zugänglich. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut.");
                return null;
            }

            bool kidNameExists = Directory.GetDirectories(groupPath).Any(dir => dir.Split(System.IO.Path.DirectorySeparatorChar).Last().Equals(kidName, StringComparison.OrdinalIgnoreCase));

            if (!kidNameExists)
            {
                _loggingService.LogAndShowError($"Kinder Name wurde nicht im Gruppen-Ordner gefunden: {kidName}", $"Der Name des Kindes wurde im Gruppenverzeichnis nicht gefunden. Bitte geben Sie einen gültigen Namen an.");
                return null;
            }

            return kidName;
        }

        public int? IsReportYearValid()
        {
            string reportYearText = _mainWindow.reportYearTextbox.Text;

            if (string.IsNullOrWhiteSpace(reportYearText))
            {
                _loggingService.LogAndShowError("Report year is empty.", "Bitte geben Sie ein gültiges Jahr ein.");
                return null;
            }

            if (!int.TryParse(reportYearText, out int parsedYear) || parsedYear < 2023 || parsedYear > 2099)
            {
                _loggingService.LogAndShowError("Das Jahr muss aus genau 4 Ziffern bestehen, und zwischen 2023 und 2099 liegen.", "Ungültiges Jahr. Bitte geben Sie ein gültiges Jahr zwischen 2023 und 2099 ein.");
                return null;
            }

            return parsedYear;
        }

        public bool AreAllRequiredFieldsFilled()
        {
            string selectedGroup = _mainWindow.groupDropdown.Text;
            string childName = _mainWindow.kidNameComboBox.Text;
            string selectedReportMonth = _mainWindow.reportMonthDropdown.Text;
            string selectedReportYear = _mainWindow.reportYearTextbox.Text;

            if (string.IsNullOrWhiteSpace(childName) || !childName.Contains(" "))
            {
                _loggingService.ShowError("Bitte geben Sie einen gültigen Namen mit Vor- und Nachnamen an.");
                return false;
            }

            if (string.IsNullOrWhiteSpace(selectedGroup) || string.IsNullOrWhiteSpace(selectedReportMonth) || string.IsNullOrWhiteSpace(selectedReportYear))
            {
                _loggingService.ShowError("Bitte füllen Sie alle geforderten Felder aus.");
                return false;
            }

            return true;
        }
    }
}