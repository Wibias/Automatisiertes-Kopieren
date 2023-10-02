﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using static Automatisiertes_Kopieren.FileManager.StringUtilities;

namespace Automatisiertes_Kopieren;

public static class ValidationHelper
{
    private static readonly LoggingService LoggingService = new();

    public static (string directoryPath, string fileName)? DetermineProtokollbogen(double monthsAndDays)
    {
        var protokollbogenRanges = new List<(double start, double end, (string directoryPath, string fileName) value)>
        {
            (10.15, 16.14,
                (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_12_Monate")),
            (16.15, 22.14,
                (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_18_Monate")),
            (22.15, 27.14,
                (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_24_Monate")),
            (27.15, 33.14,
                (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_30_Monate")),
            (33.15, 39.14,
                (Path.Combine("Entwicklungsboegen", "Krippe-Protokollboegen"), "Kind_Protokollbogen_36_Monate")),
            (39.15, 45.14,
                (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_42_Monate")),
            (45.15, 51.14,
                (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_48_Monate")),
            (51.15, 57.14,
                (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_54_Monate")),
            (57.15, 63.14,
                (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_60_Monate")),
            (63.15, 69.14,
                (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_66_Monate")),
            (69.15, 84.00, (Path.Combine("Entwicklungsboegen", "Ele-Protokollboegen"), "Kind_Protokollbogen_72_Monate"))
        };


        foreach (var (start, end, value) in protokollbogenRanges.OrderByDescending(r => r.start))
            if (monthsAndDays >= start && monthsAndDays <= end)
                return value;

        LoggingService.LogAndShowMessage($"No Protokollbogen for Month value found: {monthsAndDays}",
            $"Kein Protokollbogen für folgenden Monatswert gefunden: {monthsAndDays}",
            LoggingService.LogLevel.Warning);
        MainWindow.OperationState.OperationsSuccessful = false;
        return null;
    }

    public static double ConvertToDecimalFormat(double monthsAndDays)
    {
        var monthsAndDaysRaw = monthsAndDays.ToString("0.00", CultureInfo.InvariantCulture);
        return double.Parse(monthsAndDaysRaw.Replace(",", "."), CultureInfo.InvariantCulture);
    }

    public static string? ValidateKidName(string kidName, string homeFolder, string groupDropdownText)
    {
        if (string.IsNullOrWhiteSpace(kidName))
        {
            LoggingService.LogAndShowMessage("Kid name is empty or whitespace.",
                "Bitte geben Sie den Namen eines Kindes an.");
            return null;
        }

        var groupFolder = ConvertSpecialCharacters(groupDropdownText, ConversionType.Umlaute);

        var groupPath = $@"{homeFolder}\Entwicklungsberichte\{groupFolder} Entwicklungsberichte\Aktuell";


        if (!Directory.Exists(groupPath))
        {
            LoggingService.LogAndShowMessage($"Group path does not exist: {groupPath}",
                $"Der Pfad für den Gruppenordner {groupFolder} ist nicht zugänglich. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut.");
            return null;
        }

        var kidNameExists = Directory.GetDirectories(groupPath).Any(dir =>
            dir.Split(Path.DirectorySeparatorChar).Last().Equals(kidName, StringComparison.OrdinalIgnoreCase));

        if (kidNameExists) return kidName;
        LoggingService.LogAndShowMessage($"Kid name not found in group directory: {kidName}",
            "Der Name des Kindes wurde im Gruppenverzeichnis nicht gefunden. Bitte geben Sie einen gültigen Namen an.");
        return null;
    }

    public static int? ValidateReportYearFromTextbox(string reportYearText)
    {
        if (string.IsNullOrWhiteSpace(reportYearText)) return null;

        var isValidYear = int.TryParse(reportYearText, out var parsedYear) && parsedYear is >= 2023 and <= 2099;

        if (!isValidYear)
            throw new Exception(
                "Das Jahr muss aus genau 4 Ziffern bestehen, und zwischen 2023 und 2099 liegen. Bitte geben Sie ein gültiges Jahr ein.");

        return parsedYear;
    }
}