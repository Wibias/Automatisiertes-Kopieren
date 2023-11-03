using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Controls;
using OfficeOpenXml;
using static Automatisiertes_Kopieren.Helper.FileManagerHelper.StringUtilities;
using static Automatisiertes_Kopieren.Helper.LoggingHelper;


namespace Automatisiertes_Kopieren.Helper;

public class ExcelHelper
{
    private readonly string _homeFolder;

    public ExcelHelper(string homeFolder)
    {
        _homeFolder = homeFolder ?? throw new ArgumentNullException(nameof(homeFolder));
    }

    private string? ConvertedGroupName { get; set; }
    private string? ShortGroupName { get; set; }

    private static bool AreNamesSimilar(string name1, string name2)
    {
        const int threshold = 2;
        var distance = LevenshteinDistance(name1, name2);

        // Names are similar if the distance is within the threshold and they are not the same.
        return distance <= threshold;
    }


    private static int LevenshteinDistance(string s, string t)
    {
        var n = s.Length;
        var m = t.Length;
        var d = new int[n + 1, m + 1];

        if (n == 0) return m;
        if (m == 0) return n;

        for (var i = 0; i <= n; d[i, 0] = i++)
        for (var j = 0; j <= m; d[0, j] = j++)
        for (var x = 1; x <= n; x++)
        for (var y = 1; y <= m; y++)
        {
            var cost = t[y - 1] == s[x - 1] ? 0 : 1;
            d[x, y] = Math.Min(Math.Min(d[x - 1, y] + 1, d[x, y - 1] + 1), d[x - 1, y - 1] + cost);
        }

        return d[n, m];
    }


    public Task<(double? months, string? error, string? parsedBirthDate, string? gender)> ExtractFromExcelAsync(
        string group, string kidLastName, string kidFirstName)
    {
        string? parsedBirthDate = null;
        var genderValue = string.Empty;
        double? extractedMonths = null;

        try
        {
            ConvertedGroupName = ConvertSpecialCharacters(group, ConversionType.Umlaute);
            ShortGroupName = ConvertedGroupName.Split(' ')[0];
            var filePath =
                $@"{_homeFolder}\Entwicklungsberichte\{ConvertedGroupName} Entwicklungsberichte\Monatsrechner-Kinder-Zielsetzung-{ShortGroupName}.xlsm";

            if (string.IsNullOrEmpty(_homeFolder))
            {
                ShowMessage("Bitte setzen Sie zuerst den Heimordner.", MessageType.Error);
                return Task.FromResult<(double? months, string? error, string? parsedBirthDate, string? gender)>((null,
                    "HomeFolderNotSet", parsedBirthDate, genderValue));
            }

            try
            {
                using var package = new ExcelPackage(new FileInfo(filePath));
                var mainWorksheet = package.Workbook.Worksheets["Monatsrechner"];

                string? lastNameCell;
                string? firstNameCell;

                for (var row = 7; row <= 31; row++)
                {
                    lastNameCell = mainWorksheet.Cells[row, 3].Text;
                    firstNameCell = mainWorksheet.Cells[row, 4].Text;

                    firstNameCell = firstNameCell?.Trim();
                    lastNameCell = lastNameCell?.Trim();

                    LogMessage($"Excel name check - First Name: '{firstNameCell}', Last Name: '{lastNameCell}'",
                        LogLevel.Error);

                    var excelFirstName = firstNameCell?.Trim();
                    var excelLastName = lastNameCell?.Trim();
                    var directoryFirstName = kidFirstName.Trim();
                    var directoryLastName = kidLastName.Trim();

                    if (string.IsNullOrWhiteSpace(excelFirstName) && string.IsNullOrWhiteSpace(excelLastName)) continue;

                    if (string.Equals(excelFirstName, directoryFirstName, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(excelLastName, directoryLastName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Names are exactly the same, proceed with copying
                        var birthDate = mainWorksheet.Cells[row, 5].Text;
                        _ = DateTime.TryParse(birthDate, out var parsedDate);
                        parsedBirthDate = parsedDate.ToString("dd.MM.yyyy");

                        var monthsValueRaw = mainWorksheet.Cells[row, 6].Text;

                        if (double.TryParse(monthsValueRaw.Replace(",", "."), NumberStyles.Any,
                                CultureInfo.InvariantCulture, out var parsedValue))
                            extractedMonths = Math.Round(parsedValue, 2);
                    }
                    else if (excelLastName != null &&
                             excelFirstName != null &&
                             AreNamesSimilar(excelFirstName, directoryFirstName) &&
                             AreNamesSimilar(excelLastName, directoryLastName))
                    {
                        // Names are very similar but not the same, prompt the user to correct
                        LogAndShowMessage(
                            $"Der Name in Excel ähnelt dem Ordnernamen, ist aber nicht identisch. Excel: {excelFirstName} {excelLastName}, Ordner: {directoryFirstName} {directoryLastName}",
                            $"Der Name in Excel ähnelt dem Ordnernamen, ist aber nicht identisch.\nExcel: {excelFirstName} {excelLastName}\nOrdner: {directoryFirstName} {directoryLastName}");
                        return Task
                            .FromResult<(double? months, string? error, string? parsedBirthDate, string? gender)>((
                                null,
                                "Namen stimmen nicht überein", parsedBirthDate, genderValue));
                    }
                }

                var genderWorksheet = package.Workbook.Worksheets["NAMES-BIRTHDAYS-FILL-IN"];

                for (var row = 4; row <= 28; row++)
                {
                    lastNameCell = genderWorksheet.Cells[row, 3].Text;
                    firstNameCell = genderWorksheet.Cells[row, 4].Text;

                    if (!string.Equals(firstNameCell.Trim(), kidFirstName, StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(lastNameCell.Trim(), kidLastName, StringComparison.OrdinalIgnoreCase)) continue;

                    genderValue = genderWorksheet.Cells[row, 8].Text;
                    break;
                }
            }
            catch (FileNotFoundException)
            {
                LogAndShowMessage($"Die Datei {filePath} wurde nicht gefunden.",
                    "Die Datei wurde nicht gefunden. Bitte überprüfen Sie den Pfad.");
                return Task.FromResult<(double? months, string? error, string? parsedBirthDate, string? gender)>((null,
                    "Datei nicht gefunden", parsedBirthDate, genderValue));
            }
            catch (IOException ioEx) when (ioEx.Message.Contains("because it is being used by another process"))
            {
                LogAndShowMessage($"Die Datei {filePath} wird von einem anderen Prozess verwendet.",
                    "Die Excel-Datei ist geöffnet. Bitte schließen Sie die Datei und versuchen Sie es erneut.");
                return Task.FromResult<(double? months, string? error, string? parsedBirthDate, string? gender)>((null,
                    "Datei in Nutzung", parsedBirthDate, genderValue));
            }

            if (extractedMonths.HasValue)
                return Task.FromResult<(double? months, string? error, string? parsedBirthDate, string? gender)>((
                    extractedMonths, null, parsedBirthDate, genderValue));

            LogAndShowMessage($"Es konnte kein gültiger Monatswert für {kidFirstName} {kidLastName} extrahiert werden.",
                "Please correct the name.");
            return Task.FromResult<(double? months, string? error, string? parsedBirthDate, string? gender)>((null,
                "Extraktions Fehler", parsedBirthDate, genderValue));
        }
        catch (Exception ex)
        {
            LogAndShowMessage($"Beim Verarbeiten der Excel-Datei ist ein unerwarteter Fehler aufgetreten: {ex.Message}",
                "Ein unerwarteter Fehler ist aufgetreten. Bitte versuchen Sie es später erneut.");
            return Task.FromResult<(double? months, string? error, string? parsedBirthDate, string? gender)>((null,
                "Unerwarteter Fehler", parsedBirthDate, genderValue));
        }
    }

    private static async Task UpdateDateInWorksheetAsync(string filePath, string worksheetName, string cellAddress,
        DateTime date)
    {
        try
        {
            var fileInfo = new FileInfo(filePath);
            using var package = new ExcelPackage(fileInfo);
            var worksheet = package.Workbook.Worksheets[worksheetName];

            if (worksheet != null)
            {
                worksheet.Cells[cellAddress].Value = date;
                await package.SaveAsync();
            }
            else
            {
                throw new Exception($"The worksheet '{worksheetName}' was not found in the file {filePath}.");
            }
        }
        catch (Exception ex)
        {
            LogMessage($"Error updating worksheet '{worksheetName}': {ex.Message}", LogLevel.Error);
            throw;
        }
    }


    public async Task<bool> SelectHeutigesDatumEntwicklungsBerichtAsync(object sender, string group)
    {
        if (sender is not CheckBox { IsChecked: true }) return false;

        ConvertedGroupName = ConvertSpecialCharacters(group, ConversionType.Umlaute);
        ShortGroupName = ConvertedGroupName.Split(' ')[0];
        var filePath =
            $@"{_homeFolder}\Entwicklungsberichte\{ConvertedGroupName} Entwicklungsberichte\Monatsrechner-Kinder-Zielsetzung-{ShortGroupName}.xlsm";

        try
        {
            await UpdateDateInWorksheetAsync(filePath, "Monatsrechner", "D2", DateTime.Today);
            return true;
        }
        catch (Exception ex)
        {
            LogAndShowMessage(ex.Message, "Fehler beim Aktualisieren der Excel-Datei. Ist die Datei noch geöffnet?");
            return false;
        }
    }
}