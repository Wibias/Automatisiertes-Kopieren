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
        private readonly static LoggingService _loggingService = new LoggingService();

        public FileManager(string homeFolder)
        {
            _homeFolder = homeFolder ?? throw new ArgumentNullException(nameof(homeFolder));
        }

        public string GetTargetPath(string group, string kidName, string reportYear, string reportMonth)
        {
            group = StringUtilities.ConvertToTitleCase(group);
            group = StringUtilities.ConvertSpecialCharacters(group, StringUtilities.ConversionType.Umlaute);

            kidName = StringUtilities.ConvertToTitleCase(kidName);

            if (string.IsNullOrEmpty(_homeFolder))
            {
                throw new InvalidOperationException("Das Hauptverzeichnis ist nicht festgelegt.");
            }
            return $@"{_homeFolder}\Entwicklungsberichte\{group} Entwicklungsberichte\Aktuell\{kidName}\{reportYear}\{reportMonth}";
        }

        public void SafeRenameFile(string sourceFile, string destFile)
        {
            try
            {

                File.Move(sourceFile, destFile);
            }
            catch (Exception ex)
            {
                _loggingService.LogAndShowMessage($"Fehler beim Umbenennen der Datei: {ex.Message}", "Fehler beim Umbenennen der Datei", LoggingService.LogLevel.Error, LoggingService.MessageType.Error);
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
                _loggingService.LogMessage($"Failed to extract numeric value from protokollNumber: {protokollNumber}", LoggingService.LogLevel.Error);
                return;
            }

            string[] files = Directory.GetFiles(targetFolderPath);

            foreach (string file in files)
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                string fileExtension = Path.GetExtension(file);

                if (fileName.Equals("Allgemeiner-Entwicklungsbericht", StringComparison.OrdinalIgnoreCase) && isAllgemeinerChecked)
                {
                    string newFileName = $"{kidName}_Allgemeiner_Entwicklungsbericht_{reportMonth}_{reportYear}{fileExtension}";
                    SafeRenameFile(file, Path.Combine(targetFolderPath, newFileName));
                }

                if (fileName.Equals("Vorschul-Entwicklungsbericht", StringComparison.OrdinalIgnoreCase) && isVorschulChecked)
                {
                    string newFileName = $"{kidName}_Vorschul_Entwicklungsbericht_{reportMonth}_{reportYear}{fileExtension}";
                    SafeRenameFile(file, Path.Combine(targetFolderPath, newFileName));
                }

                if (fileName.StartsWith("Kind_Protokollbogen_", StringComparison.OrdinalIgnoreCase) && isProtokollbogenChecked)
                {
                    string newFileName = $"{kidName}_{protokollNumber}_Protokollbogen_{reportMonth}_{reportYear}{fileExtension}";
                    SafeRenameFile(file, Path.Combine(targetFolderPath, newFileName));
                }
                if (fileName.Equals("Protokoll-Elterngespraech", StringComparison.OrdinalIgnoreCase))
                {
                    string newFileName = $"{kidName}_Protokoll-Elterngespraech_{reportMonth}_{reportYear}{fileExtension}";
                    SafeRenameFile(file, Path.Combine(targetFolderPath, newFileName));
                }

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

            if (File.Exists(sourceFile))
            {
                try
                {
                    SafeCopyFile(sourceFile, Path.Combine(targetFolderPath, protokollbogenFileName));
                }
                catch (Exception ex)
                {
                    _loggingService.LogMessage($"Error copying file. Source: {sourceFile}, Destination: {Path.Combine(targetFolderPath, protokollbogenFileName)}. Error: {ex.Message}", LoggingService.LogLevel.Error);
                }
            }
            else
            {
                _loggingService.LogMessage($"File {protokollbogenFileName} not found in source folder.", LoggingService.LogLevel.Warning);
            }
        }

        public void SafeCopyFile(string sourceFile, string destFile)
        {
            try
            {
                if (File.Exists(destFile))
                {
                    bool overwrite = _loggingService.ShowMessage("Die Datei existiert bereits. Möchten Sie die vorhandene Datei überschreiben?", LoggingService.MessageType.Warning, "File exists") == MessageBoxResult.Yes;

                    if (overwrite)
                    {
                        string backupFilename = $"{Path.GetDirectoryName(destFile)}\\{DateTime.Now:yyyyMMddHHmmss}_{Path.GetFileName(destFile)}.bak";
                        File.Move(destFile, backupFilename);
                        _loggingService.ShowMessage($"Die vorhandene Datei wurde gesichert als: {backupFilename}", LoggingService.MessageType.Information, "Info");
                    }
                    else
                    {
                        _loggingService.ShowMessage("Die Datei wurde nicht kopiert.", LoggingService.MessageType.Information, "Info");
                        return;
                    }
                }

                File.Copy(sourceFile, destFile, overwrite: true);
                _loggingService.ShowMessage($"Die Datei wurde erfolgreich kopiert: {destFile}", LoggingService.MessageType.Information, "Success");
            }
            catch (Exception ex)
            {
                _loggingService.LogAndShowMessage($"Fehler beim Kopieren der Datei: {ex.Message}", "Fehler beim Kopieren der Datei", LoggingService.LogLevel.Error, LoggingService.MessageType.Error);
            }
        }

    }
}
