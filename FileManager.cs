using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using static Automatisiertes_Kopieren.LoggingService;

namespace Automatisiertes_Kopieren;

public partial class FileManager
{
    private readonly string _homeFolder;

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
            throw new InvalidOperationException("Das Hauptverzeichnis ist nicht festgelegt.");
        return
            $@"{_homeFolder}\Entwicklungsberichte\{group} Entwicklungsberichte\Aktuell\{kidName}\{reportYear}\{reportMonth}";
    }

    private static void SafeRenameFile(string sourceFile, string destFile)
    {
        try
        {
            if (File.Exists(destFile))
            {
                var result = ShowMessage("Die Datei existiert bereits. Möchtest du diese ersetzen?",
                    MessageType.Info,
                    "Confirm Replace",
                    MessageBoxButton.YesNo);

                if (result == MessageBoxResult.No) return;
                File.Delete(destFile);
            }

            File.Move(sourceFile, destFile);
        }
        catch (Exception ex)
        {
            LogAndShowMessage($"Fehler beim Umbenennen der Datei: {ex.Message}",
                "Fehler beim Umbenennen der Datei", LogLevel.Error, MessageType.Error);
        }
    }


    public static Tuple<string, string, string, string> RenameFilesInTargetDirectory(string targetFolderPath,
        string kidName,
        string reportMonth, string reportYear, bool isAllgemeinerChecked, bool isVorschuleChecked,
        bool isProtokollbogenChecked, string protokollNumber)
    {
        string? renamedProtokollbogenPath = null;
        string? renamedAllgemeinEntwicklungsberichtPath = null;
        string? renamedProtokollElterngespraechPath = null;
        string? renamedVorschuleEntwicklungsberichtPath = null;
        kidName = StringUtilities.ConvertToTitleCase(kidName);
        kidName = StringUtilities.ConvertSpecialCharacters(kidName, StringUtilities.ConversionType.Umlaute,
            StringUtilities.ConversionType.Underscore);

        reportMonth = StringUtilities.ConvertToTitleCase(reportMonth);
        reportMonth = StringUtilities.ConvertSpecialCharacters(reportMonth, StringUtilities.ConversionType.Umlaute,
            StringUtilities.ConversionType.Underscore);
        if (!int.TryParse(ProtokollNumberRegex().Match(protokollNumber).Value, out _))
        {
            LogMessage($"Failed to extract numeric value from protokollNumber: {protokollNumber}",
                LogLevel.Error);
            return new Tuple<string, string, string, string>(renamedProtokollbogenPath ?? string.Empty,
                renamedAllgemeinEntwicklungsberichtPath ?? string.Empty,
                renamedProtokollElterngespraechPath ?? string.Empty,
                renamedVorschuleEntwicklungsberichtPath ?? string.Empty);
        }


        var files = Directory.GetFiles(targetFolderPath);

        foreach (var file in files)
        {
            var fileName = Path.GetFileNameWithoutExtension(file);
            var fileExtension = Path.GetExtension(file);

            if (fileName.Equals("Allgemeiner-Entwicklungsbericht", StringComparison.OrdinalIgnoreCase) &&
                isAllgemeinerChecked)
            {
                var newFileName =
                    $"{kidName}_Allgemeiner_Entwicklungsbericht_{reportMonth}_{reportYear}{fileExtension}";
                SafeRenameFile(file, Path.Combine(targetFolderPath, newFileName));
                renamedAllgemeinEntwicklungsberichtPath = Path.Combine(targetFolderPath, newFileName);
            }

            if (fileName.Equals("Vorschule-Entwicklungsbericht", StringComparison.OrdinalIgnoreCase) &&
                isVorschuleChecked)
            {
                var newFileName = $"{kidName}_Vorschule_Entwicklungsbericht_{reportMonth}_{reportYear}{fileExtension}";
                SafeRenameFile(file, Path.Combine(targetFolderPath, newFileName));
                renamedVorschuleEntwicklungsberichtPath = Path.Combine(targetFolderPath, newFileName);
            }


            if (fileName.StartsWith("Kind_Protokollbogen_", StringComparison.OrdinalIgnoreCase) &&
                isProtokollbogenChecked)
            {
                var newFileName =
                    $"{kidName}_{protokollNumber}_Protokollbogen_{reportMonth}_{reportYear}{fileExtension}";
                SafeRenameFile(file, Path.Combine(targetFolderPath, newFileName));
                renamedProtokollbogenPath = Path.Combine(targetFolderPath, newFileName);
            }

            if (!fileName.Equals("Protokoll-Elterngespraech", StringComparison.OrdinalIgnoreCase)) continue;
            {
                var newFileName = $"{kidName}_Protokoll_Elterngespraech_{reportMonth}_{reportYear}{fileExtension}";
                SafeRenameFile(file, Path.Combine(targetFolderPath, newFileName));
                renamedProtokollElterngespraechPath = Path.Combine(targetFolderPath, newFileName);
            }
        }

        return new Tuple<string, string, string, string>(renamedProtokollbogenPath ?? string.Empty,
            renamedAllgemeinEntwicklungsberichtPath ?? string.Empty,
            renamedProtokollElterngespraechPath ?? string.Empty,
            renamedVorschuleEntwicklungsberichtPath ?? string.Empty);
    }

    public static void CopyFilesFromSourceToTarget(string? sourceFile, string targetFolderPath,
        string protokollbogenFileName)
    {
        if (!Directory.Exists(targetFolderPath)) Directory.CreateDirectory(targetFolderPath);

        if (sourceFile != null && File.Exists(sourceFile))
            try
            {
                SafeCopyFile(sourceFile, Path.Combine(targetFolderPath, protokollbogenFileName));
            }
            catch (Exception ex)
            {
                LogMessage(
                    $"Error copying file. Source: {sourceFile}, Destination: {Path.Combine(targetFolderPath, protokollbogenFileName)}. Error: {ex.Message}",
                    LogLevel.Error);
            }
        else
            LogMessage($"File {protokollbogenFileName} not found in source folder.", LogLevel.Warning);
    }

    private static void SafeCopyFile(string sourceFile, string destFile)
    {
        try
        {
            if (File.Exists(destFile))
            {
                var result = ShowMessage("Möchten Sie das Hauptverzeichnis ändern?", MessageType.Info,
                    "Hauptverzeichnis nicht festgelegt", MessageBoxButton.YesNo);

                if (result == MessageBoxResult.Yes)
                {
                    var backupFilename =
                        $"{Path.GetDirectoryName(destFile)}\\{DateTime.Now:yyyyMMddHHmmss}_{Path.GetFileName(destFile)}.bak";
                    File.Move(destFile, backupFilename);
                    ShowMessage($"Die vorhandene Datei wurde gesichert als: {backupFilename}",
                        MessageType.Info, "Info");
                }
                else
                {
                    ShowMessage("Die Datei wurde nicht kopiert.", MessageType.Info, "Info");
                    return;
                }
            }

            File.Copy(sourceFile, destFile, true);
        }
        catch (Exception ex)
        {
            LogAndShowMessage($"Fehler beim Kopieren der Datei: {ex.Message}",
                "Fehler beim Kopieren der Datei", LogLevel.Error, MessageType.Error);
        }
    }

    [GeneratedRegex("\\d+")]
    private static partial Regex ProtokollNumberRegex();

    public static class StringUtilities
    {
        public enum ConversionType
        {
            Umlaute,
            Underscore
        }

        public static string ConvertToTitleCase(string inputString)
        {
            if (string.IsNullOrWhiteSpace(inputString))
                return string.Empty;

            var textInfo = new CultureInfo("de-DE", false).TextInfo;
            return textInfo.ToTitleCase(inputString.ToLower());
        }

        public static string ConvertSpecialCharacters(string input, params ConversionType[] types)
        {
            return types.Aggregate(input, (current, type) => type switch
            {
                ConversionType.Umlaute => current.Replace("ä", "ae").Replace("ö", "oe"),
                ConversionType.Underscore => current.Replace(" ", "_"),
                _ => current
            });
        }
    }
}