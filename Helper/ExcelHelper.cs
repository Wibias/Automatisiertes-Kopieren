using OfficeOpenXml;
using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Controls;
using static Automatisiertes_Kopieren.Helper.FileManagerHelper.StringUtilities;
using static Automatisiertes_Kopieren.Helper.LoggingHelper;

namespace Automatisiertes_Kopieren.Helper
{
    public class ExcelHelper
    {
        private readonly string _homeFolder;

        public ExcelHelper(string homeFolder)
        {
            _homeFolder = homeFolder ?? throw new ArgumentNullException(nameof(homeFolder));
        }

        private string? ConvertedGroupName { get; set; }
        private string? ShortGroupName { get; set; }

        public async Task<(double? months, string? error, string? parsedBirthDate, string? gender)> ExtractFromExcelAsync(
            string group, string kidLastName, string kidFirstName)
        {
            return await Task.Run<(double?, string?, string?, string?)>(() =>
            {
                var parsedBirthDate = (string?)null;
                var genderValue = string.Empty;
                var extractedMonths = (double?)null;
                var compatibleFilePath = (string?)null;
                string originalFilePath = Path.Combine(_homeFolder, $"Entwicklungsberichte\\{ConvertedGroupName} Entwicklungsberichte\\Monatsrechner-Kinder-Zielsetzung-{ShortGroupName}.xlsm");


                if (string.IsNullOrEmpty(_homeFolder))
                {
                    ShowMessage("Bitte setzen Sie zuerst den Heimordner.", MessageType.Error);
                    return (null, "HomeFolderNotSet", parsedBirthDate, genderValue);
                }

                try
                {
                    ConvertedGroupName = ConvertSpecialCharacters(group, ConversionType.Umlaute);
                    ShortGroupName = ConvertedGroupName.Split(' ')[0];
                    originalFilePath = Path.Combine(_homeFolder, $"Entwicklungsberichte\\{ConvertedGroupName} Entwicklungsberichte\\Monatsrechner-Kinder-Zielsetzung-{ShortGroupName}.xlsm");

                    compatibleFilePath = Path.ChangeExtension(originalFilePath, ".xlsx");

                    ConvertXlsxForCompatibility(originalFilePath, compatibleFilePath);

                    using (var package = new ExcelPackage(new FileInfo(compatibleFilePath)))
                    {
                        extractedMonths = ExtractMonths(package, kidFirstName, kidLastName, out parsedBirthDate);
                        genderValue = ExtractGender(package, kidFirstName, kidLastName);

                        if (extractedMonths.HasValue)
                        {
                            return (extractedMonths, null, parsedBirthDate, genderValue);
                        }

                    LogAndShowMessage($"Es konnte kein gültiger Monatswert für {kidFirstName} {kidLastName} extrahiert werden.", "Bitte korrigieren Sie den Namen.");
                    return (null, "Extraktions Fehler", parsedBirthDate, genderValue);
                    }
                }
                catch (FileNotFoundException ex)
                {
                    LogException(ex, $"Die Datei '{originalFilePath}' wurde nicht gefunden.");
                    return (null, "Datei nicht gefunden", parsedBirthDate, genderValue);
                }
                catch (IOException ioEx)
                {
                    if (ioEx.Message.Contains("because it is being used by another process"))
                    {
                        LogException(ioEx, $"Die Datei '{originalFilePath}' wird von einem anderen Prozess verwendet.");
                        return (null, "Datei in Nutzung", parsedBirthDate, genderValue);
                    }
                    else
                    {
                        LogException(ioEx, $"Ein Fehler ist aufgetreten beim Zugriff auf die Datei '{originalFilePath}': {ioEx.Message}");
                        return (null, "IO-Fehler", parsedBirthDate, genderValue);
                    }
                }
                catch (Exception ex)
                {
                    LogException(ex, $"Ein unerwarteter Fehler ist aufgetreten: {ex.Message}");
                    return (null, "Unerwarteter Fehler", parsedBirthDate, genderValue);
                }
                finally
                {
                    if (!string.IsNullOrEmpty(compatibleFilePath) && File.Exists(compatibleFilePath))
                    {
                        try
                        {
                            File.Delete(compatibleFilePath);
                        }
                        catch (Exception ex)
                        {
                            LogAndShowMessage($"Fehler beim Löschen der temporären .xlsx Datei: {ex.Message}", "Ein Fehler ist aufgetreten beim Löschen der temporären .xlsx Datei.");
                        }
                    }
                }
            });
        }

        private double? ExtractMonths(ExcelPackage package, string kidFirstName, string kidLastName, out string? parsedBirthDate)
        {
            var mainWorksheet = package.Workbook.Worksheets["Monatsrechner"];
            parsedBirthDate = null;

            for (var row = 7; row <= 31; row++)
            {
                string? lastNameCell = mainWorksheet.Cells[row, 3].Text?.Trim();
                string? firstNameCell = mainWorksheet.Cells[row, 4].Text?.Trim();

                if (string.IsNullOrWhiteSpace(firstNameCell) && string.IsNullOrWhiteSpace(lastNameCell)) continue;

                if (NamesMatch(firstNameCell, lastNameCell, kidFirstName, kidLastName))
                {
                    var birthDate = mainWorksheet.Cells[row, 5].Text;
                    if (DateTime.TryParse(birthDate, out var parsedDate))
                    {
                        parsedBirthDate = parsedDate.ToString("dd.MM.yyyy");

                        var monthsValueRaw = mainWorksheet.Cells[row, 6].Text;
                        if (double.TryParse(monthsValueRaw.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedValue))
                        {
                            return Math.Round(parsedValue, 2);
                        }
                    }
                }
                else if (NamesSimilar(firstNameCell, lastNameCell, kidFirstName, kidLastName))
                {
                    LogAndShowMessage($"Der Name in Excel ähnelt dem Ordnernamen, ist aber nicht identisch. Excel: {firstNameCell} {lastNameCell}, Ordner: {kidFirstName} {kidLastName}",
                        $"Der Name in Excel ähnelt dem Ordnernamen, ist aber nicht identisch.\nExcel: {firstNameCell} {lastNameCell}\nOrdner: {kidFirstName} {kidLastName}");
                    return null;
                }
            }
            return null;
        }

        private string ExtractGender(ExcelPackage package, string kidFirstName, string kidLastName)
        {
            var genderWorksheet = package.Workbook.Worksheets["NAMES-BIRTHDAYS-FILL-IN"];

            for (var row = 4; row <= 28; row++)
            {
                string? lastNameCell = genderWorksheet.Cells[row, 3].Text?.Trim();
                string? firstNameCell = genderWorksheet.Cells[row, 4].Text?.Trim();

                if (string.Equals(firstNameCell, kidFirstName, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(lastNameCell, kidLastName, StringComparison.OrdinalIgnoreCase))
                {
                    return genderWorksheet.Cells[row, 8].Text;
                }
            }
            return string.Empty;
        }

        private bool NamesMatch(string? excelFirstName, string? excelLastName, string directoryFirstName, string directoryLastName)
        {
            return string.Equals(excelFirstName, directoryFirstName, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(excelLastName, directoryLastName, StringComparison.OrdinalIgnoreCase);
        }

        private bool NamesSimilar(string? excelFirstName, string? excelLastName, string directoryFirstName, string directoryLastName)
        {
            return excelFirstName != null && excelLastName != null &&
                   StringHelpers.AreNamesSimilar(excelFirstName, directoryFirstName) &&
                   StringHelpers.AreNamesSimilar(excelLastName, directoryLastName);
        }

        private static void ConvertXlsxForCompatibility(string inputFilePath, string outputFilePath)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(inputFilePath)))
                {
                    package.SaveAs(new FileInfo(outputFilePath));
                }
            }
            catch (Exception ex)
            {
                LogAndShowMessage($"An error occurred while converting {inputFilePath} to {outputFilePath}: {ex.Message}", "Error");
                throw;
            }
        }

        private static async Task UpdateDateInWorksheetAsync(string filePath, string worksheetName, string cellAddress, DateTime date)
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
                    throw new Exception($"Das Arbeitsblatt '{worksheetName}' wurde nicht in der Datei {filePath} gefunden.");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Fehler beim Aktualisieren des Arbeitsblatts '{worksheetName}': {ex.Message}", LogLevel.Error);
                throw;
            }
        }

        public async Task<bool> SelectHeutigesDatumEntwicklungsBerichtAsync(object sender, string group)
        {
            if (sender is not CheckBox { IsChecked: true }) return false;

            ConvertedGroupName = ConvertSpecialCharacters(group, ConversionType.Umlaute);
            ShortGroupName = ConvertedGroupName.Split(' ')[0];
            var filePath = Path.Combine
                (_homeFolder, "Entwicklungsberichte", $"{ConvertedGroupName} Entwicklungsberichte", $"Monatsrechner-Kinder-Zielsetzung-{ShortGroupName}.xlsm");

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
}
