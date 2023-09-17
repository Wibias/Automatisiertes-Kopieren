using System;
using System.Globalization;
using System.IO;
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

        public string GetTargetPath(string group, string kidName, string reportMonth, string reportYear)
        {
            group = StringUtilities.ConvertToTitleCase(group);
            kidName = StringUtilities.ConvertToTitleCase(kidName);
            reportMonth = StringUtilities.ConvertToTitleCase(reportMonth);

            if (string.IsNullOrEmpty(_homeFolder))
            {
                throw new InvalidOperationException("Home folder is not set.");
            }
            return $@"{_homeFolder}\Entwicklungsberichte\{group}\{kidName}\{reportYear}\{reportMonth}";
        }

        public void RenameFilesInTargetDirectory(string targetFolderPath, string kidName, string reportMonth, string reportYear, bool isAllgemeinerChecked, bool isVorschulChecked, bool isProtokollbogenChecked, int protokollNumber)
        {
            kidName = StringUtilities.ConvertToTitleCase(kidName);
            kidName = StringUtilities.ConvertSpecialCharacters(kidName, StringUtilities.ConversionType.Umlaute, StringUtilities.ConversionType.Underscore);

            reportMonth = StringUtilities.ConvertToTitleCase(reportMonth);
            reportMonth = StringUtilities.ConvertSpecialCharacters(reportMonth, StringUtilities.ConversionType.Umlaute, StringUtilities.ConversionType.Underscore);

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
                    string newFileName = $"Kind_Protokollbogen_{protokollNumber}_Monate{fileExtension}";
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
                MessageBox.Show($"Error copying file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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

        public string? GetSourceFolder(string protokollbogen)
        {
            string? homeFolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            // ... Rest of the logic ...

            return null;
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

        public void CopyFilesFromSourceToTarget(string sourceFolderPath, string targetFolderPath)
        {
            if (!Directory.Exists(targetFolderPath))
            {
                Directory.CreateDirectory(targetFolderPath);
            }

            foreach (string file in Directory.GetFiles(sourceFolderPath))
            {
                string fileName = Path.GetFileName(file);
                SafeCopyFile(file, Path.Combine(targetFolderPath, fileName));
            }
        }
    }
 }
