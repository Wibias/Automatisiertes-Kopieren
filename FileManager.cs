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
            group = StringUtilities.ConvertSpecialCharacters(group, StringUtilities.ConversionType.Umlaute); // Convert special characters for the group name

            kidName = StringUtilities.ConvertToTitleCase(kidName);

            if (string.IsNullOrEmpty(_homeFolder))
            {
                throw new InvalidOperationException("Das Hauptverzeichnis ist nicht festgelegt.");
            }
            return $@"{_homeFolder}\Entwicklungsberichte\{group} Entwicklungsberichte\Aktuell\{kidName}\{reportYear}";
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
                // Handle the error if the number couldn't be extracted.
                Log.Error($"Failed to extract numeric value from protokollNumber: {protokollNumber}");
                return; // or throw an exception, or handle it in another appropriate way.
            }

            string[] files = Directory.GetFiles(targetFolderPath);

            foreach (string file in files)
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                string fileExtension = Path.GetExtension(file);

                if (fileName.Equals("Allgemeiner Entwicklungsbericht", StringComparison.OrdinalIgnoreCase) && isAllgemeinerChecked)
                {
                    string newFileName = $"{kidName}_Entwicklungsbericht_Allgemein_{reportMonth}_{reportYear}{fileExtension}";
                    File.Move(file, Path.Combine(targetFolderPath, newFileName));
                }

                if (fileName.Equals("Vorschulentwicklungsbericht", StringComparison.OrdinalIgnoreCase) && isVorschulChecked)
                {
                    string newFileName = $"{kidName}_Vorschulentwicklungsbericht_{reportMonth}_{reportYear}{fileExtension}";
                    File.Move(file, Path.Combine(targetFolderPath, newFileName));
                }

                if (fileName.StartsWith("Kind_Protokollbogen_", StringComparison.OrdinalIgnoreCase) && isProtokollbogenChecked)
                {
                    string newFileName = $"{kidName}_{protokollNumber}_Monate_{reportMonth}_{reportYear}{fileExtension}";
                    File.Move(file, Path.Combine(targetFolderPath, newFileName));
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
                MessageBox.Show($"Error beim Kopieren der Datei: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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

        public void CopyFilesFromSourceToTarget(string sourceFile, string targetFolderPath, string protokollbogenFileName)
        {
            if (!Directory.Exists(targetFolderPath))
            {
                Directory.CreateDirectory(targetFolderPath);
            }

            Serilog.Log.Information($"Trying to access file at: {sourceFile}");
            if (File.Exists(sourceFile))
            {
                try
                {
                    SafeCopyFile(sourceFile, Path.Combine(targetFolderPath, protokollbogenFileName));
                }
                catch (Exception ex)
                {
                    Serilog.Log.Error($"Error copying file. Source: {sourceFile}, Destination: {Path.Combine(targetFolderPath, protokollbogenFileName)}. Error: {ex.Message}");
                }
            }
            else
            {
                Serilog.Log.Warning($"File {protokollbogenFileName} not found in source folder.");
            }
        }

        public void SafeCopyFile(string sourceFile, string destFile)
        {
            try
            {
                // Check if destination file exists
                if (File.Exists(destFile))
                {
                    // Prompt user to overwrite or not
                    MessageBoxResult result = MessageBox.Show("Die Datei existiert bereits. Möchten Sie die vorhandene Datei überschreiben?", "File exists", MessageBoxButton.YesNo, MessageBoxImage.Question);

                    if (result == MessageBoxResult.Yes)
                    {
                        // Backup existing file with timestamp
                        string backupFilename = $"{Path.GetDirectoryName(destFile)}\\{DateTime.Now:yyyyMMddHHmmss}_{Path.GetFileName(destFile)}.bak";
                        File.Move(destFile, backupFilename);
                        MessageBox.Show($"Die vorhandene Datei wurde gesichert als: {backupFilename}", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("Die Datei wurde nicht kopiert.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }
                }

                // Copy source file to destination
                File.Copy(sourceFile, destFile, overwrite: true);
                MessageBox.Show($"Die Datei wurde erfolgreich kopiert: {destFile}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Kopieren der Datei: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

    }
}
