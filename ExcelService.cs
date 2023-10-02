using ClosedXML.Excel;
using OfficeOpenXml;
using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using static System.DateTime;
using static Automatisiertes_Kopieren.FileManager.StringUtilities;
using static Automatisiertes_Kopieren.LoggingService;

namespace Automatisiertes_Kopieren;

public class ExcelService
{
    private readonly string _homeFolder;
    private static readonly LoggingService LoggingService = new();

    public ExcelService(string homeFolder)
    {
        _homeFolder = homeFolder ?? throw new ArgumentNullException(nameof(homeFolder));
    }
    private string? ConvertedGroupName { get; set; }
    private string? ShortGroupName { get; set; }
    public (double? months, string? error, string? parsedBirthDate, string? gender) ExtractFromExcel(string group, string kidLastName, string kidFirstName)
    {
        ConvertedGroupName = ConvertSpecialCharacters(group, ConversionType.Umlaute);
        ShortGroupName = ConvertedGroupName.Split(' ')[0];
        var filePath = $@"{_homeFolder}\Entwicklungsberichte\{ConvertedGroupName} Entwicklungsberichte\Monatsrechner-Kinder-Zielsetzung-{ShortGroupName}.xlsm";
        string? parsedBirthDate = null;
        var genderValue = string.Empty;
        double? extractedMonths = null;

        if (string.IsNullOrEmpty(_homeFolder))
        {
            LoggingService.ShowMessage("Bitte setzen Sie zuerst den Heimordner.", MessageType.Error);
            return (null, "HomeFolderNotSet", parsedBirthDate, genderValue);
        }

        try
        {
            using var workbook = new XLWorkbook(filePath);
            var mainWorksheet = workbook.Worksheet("Monatsrechner");

            for (var row = 7; row <= 31; row++)
            {
                var lastNameCell = mainWorksheet.Cell(row, 3).Value.ToString();
                var firstNameCell = mainWorksheet.Cell(row, 4).Value.ToString();

                if (lastNameCell != null)
                {
                    lastNameCell = lastNameCell.Trim();
                    if (firstNameCell != null)
                    {
                        firstNameCell = firstNameCell.Trim();

                        if (!string.Equals(lastNameCell, kidLastName, StringComparison.OrdinalIgnoreCase) ||
                            !string.Equals(firstNameCell, kidFirstName, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                }

                var birthDate = mainWorksheet.Cell(row, 5).Value.ToString();
                var parseBirthDate = TryParse(birthDate, out var parsedDate);
                parsedBirthDate = parsedDate.ToString("dd.MM.yyyy");

                var monthsValueRaw = mainWorksheet.Cell(row, 6).Value.ToString();

                if (monthsValueRaw != null && double.TryParse(monthsValueRaw.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedValue))
                {
                    extractedMonths = Math.Round(parsedValue, 2);
                }
            }

            var genderWorksheet = workbook.Worksheet("NAMES-BIRTHDAYS-FILL-IN");

            for (var row = 4; row <= 28; row++)
            {
                var lastNameCell = genderWorksheet.Cell(row, 3).Value.ToString();
                var firstNameCell = genderWorksheet.Cell(row, 4).Value.ToString();

                if (lastNameCell != null)
                {
                    lastNameCell = lastNameCell.Trim();
                    if (firstNameCell != null)
                    {
                        firstNameCell = firstNameCell.Trim();

                        if (!string.Equals(lastNameCell, kidLastName, StringComparison.OrdinalIgnoreCase) ||
                            !string.Equals(firstNameCell, kidFirstName, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                }

                genderValue = genderWorksheet.Cell(row, 8).Value.ToString();
                break;
            }
        }
        catch (FileNotFoundException)
        {
            LoggingService.LogAndShowMessage($"Die Datei {filePath} wurde nicht gefunden.",
                "Die Datei wurde nicht gefunden. Bitte überprüfen Sie den Pfad.");
            return (null, "FileNotFound", parsedBirthDate, genderValue);
        }
        catch (IOException ioEx) when (ioEx.Message.Contains("because it is being used by another process"))
        {
            LoggingService.LogAndShowMessage($"Die Datei {filePath} wird von einem anderen Prozess verwendet.",
                "Die Excel-Datei ist geöffnet. Bitte schließen Sie die Datei und versuchen Sie es erneut.");
            return (null, "FileInUse", parsedBirthDate, genderValue);
        }
        catch (Exception ex)
        {
            LoggingService.LogAndShowMessage($"Beim Verarbeiten der Excel-Datei ist ein unerwarteter Fehler aufgetreten: {ex.Message}",
                "Ein unerwarteter Fehler ist aufgetreten. Bitte versuchen Sie es später erneut.");
            return (null, "UnexpectedError", parsedBirthDate, genderValue);
        }

        if (extractedMonths.HasValue) return (extractedMonths, null, parsedBirthDate, genderValue);
        LoggingService.LogAndShowMessage($"Es konnte kein gültiger Monatswert für {kidFirstName} {kidLastName} extrahiert werden.",
            "Es konnte kein gültiger Monatswert extrahiert werden. Bitte überprüfen Sie die Daten.");
        return (null, "ExtractionError", parsedBirthDate, genderValue);

    }

    private static void UpdateDateInWorksheet(string filePath, string worksheetName, string cellAddress, DateTime date)
    {
        var fileInfo = new FileInfo(filePath);
        using var package = new ExcelPackage(fileInfo);
        var worksheet = package.Workbook.Worksheets[worksheetName];

        if (worksheet == null)
        {
            throw new Exception($"The worksheet '{worksheetName}' was not found in the file {filePath}.");
        }

        worksheet.Cells[cellAddress].Value = date;
        package.Save();
    }
    public string GetExcelFilePath(string groupName)
    {
        var convertedGroupName = ConvertSpecialCharacters(groupName, ConversionType.Umlaute);
        var shortGroupName = convertedGroupName.Split(' ')[0];
        return $@"{_homeFolder}\Entwicklungsberichte\{convertedGroupName} Entwicklungsberichte\Monatsrechner-Kinder-Zielsetzung-{shortGroupName}.xlsm";
    }

    public void SelectHeutigesDatumEntwicklungsBericht(object sender, RoutedEventArgs e)
    {
        if (sender is not CheckBox { IsChecked: true }) return;
        var filePath = $@"{_homeFolder}\Entwicklungsberichte\{ConvertedGroupName} Entwicklungsberichte\Monatsrechner-Kinder-Zielsetzung-{ShortGroupName}.xlsm";

        try
        {
            UpdateDateInWorksheet(filePath, "Monatsrechner", "D2", Today);
        }
        catch (Exception ex)
        {
            LoggingService.LogAndShowMessage(ex.Message, "Error updating Excel file.");
        }
    }
}