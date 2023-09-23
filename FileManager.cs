using Serilog;
using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;

namespace Automatisiertes_Kopieren
{
    public class FileManager
    {
        private readonly string _homeFolder;

        public FileManager(string homeFolder)
        {
            _homeFolder = homeFolder ?? throw new ArgumentNullException(nameof(homeFolder));
        }

        public string GetTargetPath(string group, string kidName, string reportYear)
        {
            group = StringUtilities.ConvertToTitleCase(group);
            group = StringUtilities.ConvertSpecialCharacters(group, StringUtilities.ConversionType.Umlaute);

            kidName = StringUtilities.ConvertToTitleCase(kidName);

            if (string.IsNullOrEmpty(_homeFolder))
            {
                throw new InvalidOperationException("Das Hauptverzeichnis ist nicht festgelegt.");
            }
            return Path.Combine(_homeFolder, "Entwicklungsberichte", $"{group} Entwicklungsberichte", "Aktuell", kidName, reportYear);
        }

        public void SafeRenameFile(string sourceFile, string destFile)
        {
            try
            {
                // Check if destination file exists
                if (File.Exists(destFile))
                {
                    // Prompt user to overwrite or not
                    MessageBoxResult result = MessageBox.Show("Die Datei existiert bereits. Möchten Sie die vorhandene Datei überschreiben?", "Datei existiert", MessageBoxButton.YesNo, MessageBoxImage.Question);

                    if (result == MessageBoxResult.Yes)
                    {
                        // Backup existing file with timestamp
                        string backupFilename = $"{Path.GetDirectoryName(destFile)}\\{DateTime.Now:yyyyMMddHHmmss}_{Path.GetFileName(destFile)}.bak";
                        File.Move(destFile, backupFilename);
                        MessageBox.Show($"Die vorhandene Datei wurde gesichert als: {backupFilename}", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("Die Datei wurde nicht umbenannt.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }
                }

                File.Move(sourceFile, destFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Umbenennen der Datei: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        public void RenameFilesInTargetDirectory(string targetFolderPath, string kidName, string reportMonth, string reportYear, bool isAllgemeinerChecked, bool isVorschulChecked, bool isProtokollbogenChecked, string protokollNumber)
        {
            kidName = StringUtilities.ConvertToTitleCase(kidName);
            kidName = StringUtilities.ConvertSpecialCharacters(kidName, StringUtilities.ConversionType.Umlaute, StringUtilities.ConversionType.Underscore);

            reportMonth = StringUtilities.ConvertToTitleCase(reportMonth);
            reportMonth = StringUtilities.ConvertSpecialCharacters(reportMonth, StringUtilities.ConversionType.Umlaute, StringUtilities.ConversionType.Underscore);
            int numericProtokollNumber;
            if (int.TryParse(Regex.Match(protokollNumber, @"\d+").Value, out numericProtokollNumber))
            {
                // Now you can use numericProtokollNumber as an integer.
            }
            else
            {
                Log.Error($"Der numerische Wert konnte nicht aus folgender Protokollnummer extrahiert werden: {protokollNumber}");
                return;
            }

            string[] files = Directory.GetFiles(targetFolderPath);

            foreach (string file in files)
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                string fileExtension = Path.GetExtension(file);

                if (fileName.Equals("Allgemeiner-Entwicklungsbericht", StringComparison.OrdinalIgnoreCase) && isAllgemeinerChecked)
                {
                    string newFileName = $"{kidName}_Allgemeiner-Entwicklungsbericht_{reportMonth}_{reportYear}{fileExtension}";
                    SafeRenameFile(file, Path.Combine(targetFolderPath, newFileName));
                }

                if (fileName.Equals("Vorschul-Entwicklungsbericht", StringComparison.OrdinalIgnoreCase) && isVorschulChecked)
                {
                    string newFileName = $"{kidName}_Vorschul-Entwicklungsbericht_{reportMonth}_{reportYear}{fileExtension}";
                    SafeRenameFile(file, Path.Combine(targetFolderPath, newFileName));
                }

                if (fileName.StartsWith("Kind_Protokollbogen_", StringComparison.OrdinalIgnoreCase) && isProtokollbogenChecked)
                {
                    string newFileName = $"{kidName}_{protokollNumber}_Monate_{reportMonth}_{reportYear}{fileExtension}";
                    SafeRenameFile(file, Path.Combine(targetFolderPath, newFileName));
                }
            }
        }

        public void CopyFile(string sourcePath, string targetPath)
        {
            try
            {
                File.Copy(sourcePath, targetPath, overwrite: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Kopieren der Datei: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public static class StringUtilities
        {
            public static string ConvertToTitleCase(string inputString)
            {
                if (string.IsNullOrWhiteSpace(inputString))
                    return string.Empty;

                TextInfo textInfo = new CultureInfo("de-DE", false).TextInfo;
                return textInfo.ToTitleCase(inputString.ToLower());
            }

            public static string ConvertSpecialCharacters(string input, params ConversionType[] types)
            {
                foreach (var type in types)
                {
                    switch (type)
                    {
                        case ConversionType.Umlaute:
                            input = input.Replace("ä", "ae").Replace("ö", "oe");
                            break;
                        case ConversionType.Underscore:
                            input = input.Replace(" ", "_");
                            break;
                    }
                }
                return input;
            }

            public enum ConversionType
            {
                Umlaute,
                Underscore
            }
        }

        public void SafeCopyFile(string sourceFile, string destFile)
        {
            string? destDir = Path.GetDirectoryName(destFile);
            if (destDir != null && !Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            Serilog.Log.Information($"Versuche, auf die Datei zuzugreifen unter: {sourceFile}");

            if (!File.Exists(sourceFile))
            {
                Serilog.Log.Warning($"Datei {Path.GetFileName(sourceFile)} wurde nicht im Quellverzeichnis gefunden.");
                return;
            }

            if (destDir == null || !ValidationHelper.IsValidPath(destDir))
            {
                Log.Error($"Der Gruppenpfad ist nicht gültig oder zugänglich: {destDir ?? "null"}");
                MessageBox.Show($"Der Zielordner ist nicht gültig oder zugänglich. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut.", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                if (File.Exists(destFile))
                {
                    MessageBoxResult result = MessageBox.Show("Die Datei existiert bereits. Möchten Sie die vorhandene Datei überschreiben?", "Datei existiert bereits", MessageBoxButton.YesNo, MessageBoxImage.Question);

                    if (result == MessageBoxResult.Yes)
                    {
                        string backupFilename = $"{Path.GetDirectoryName(destFile)}\\{DateTime.Now:yyyyMMddHHmmss}_{Path.GetFileName(destFile)}.bak";
                        File.Move(destFile, backupFilename);
                        MessageBox.Show($"Die vorhandene Datei wurde gesichert als: {backupFilename}", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("Die Datei wurde nicht kopiert.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }
                }

                File.Copy(sourceFile, destFile, overwrite: true);
                MessageBox.Show($"Die Datei wurde erfolgreich kopiert: {destFile}", "Erfolgreich", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Kopieren der Datei: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
