using System;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Automatisiertes_Kopieren.Helper;
using AutoUpdaterDotNET;
using Ookii.Dialogs.Wpf;
using Serilog;
using static Automatisiertes_Kopieren.Helper.FileManagerHelper.StringUtilities;
using static Automatisiertes_Kopieren.Helper.PdfHelper;
using static Automatisiertes_Kopieren.Helper.LoggingHelper;
using Size = System.Drawing.Size;

namespace Automatisiertes_Kopieren;

public partial class MainWindow
{
    private readonly AutoCompleteHelper _autoCompleteHelper;
    private readonly ExcelHelper _excelHelper;
    private FileManagerHelper? _fileManager;

    private string? _homeFolder;
    private bool _isHandlingCheckboxEvent;

    public MainWindow()
    {
        InitializeLogger();
        _autoCompleteHelper = new AutoCompleteHelper(this);
        InitializeComponent();

        AutoUpdater.Start("https://raw.githubusercontent.com/enkama/Automatisiertes-Kopieren/main/autoupdater.xml");
        AutoUpdater.UpdateFormSize = new Size(800, 600);

        var settings = Settings.LoadSettings();
        if (settings != null && !string.IsNullOrEmpty(settings.HomeFolderPath))
        {
            HomeFolder = settings.HomeFolderPath;
        }
        else
        {
            SelectHomeFolder();
            if (string.IsNullOrEmpty(HomeFolder))
            {
                ShowMessage("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", MessageType.Error);
                throw new InvalidOperationException("Hauptverzeichnis muss gesetzt sein.");
            }
        }

        _excelHelper = new ExcelHelper(HomeFolder);
        ProtokollbogenAutoCheckbox.Checked += OnProtokollbogenAutoCheckboxChanged;
        ProtokollbogenAutoCheckbox.Unchecked += OnProtokollbogenAutoCheckboxChanged;
    }

    public string? HomeFolder
    {
        get => _homeFolder;
        private set => _ = SetHomeFolderAsync(value);
    }

    private async Task SetHomeFolderAsync(string? value)
    {
        _homeFolder = value;
        if (GroupDropdown.SelectedIndex != 0 || string.IsNullOrEmpty(_homeFolder)) return;
        var defaultKidNames = await _autoCompleteHelper.GetKidNamesForGroupAsync("Bären");
        KidNameComboBox.ItemsSource = defaultKidNames;
    }

    public void SetHomeFolder(string path)
    {
        HomeFolder = path;
    }

    private void OnSelectHomeFolderButtonClicked(object sender, RoutedEventArgs e)
    {
        SelectHomeFolder();
    }

    private void InitializeFileManager()
    {
        if (HomeFolder != null)
            _fileManager = new FileManagerHelper(HomeFolder);
        else
            ShowMessage("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", MessageType.Error);
    }

    private async void OnSelectHeutigesDatumEntwicklungsBericht(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(KidNameComboBox.Text))
        {
            ShowMessage("Bitte wählen Sie einen Kindernamen, bevor Sie das Arbeitsblatt aktualisieren.", MessageType.Error);
            UpdateMonatsrechnerCheckBox.IsChecked = false;
            return;
        }

        var group = GroupDropdown.Text;
        var success =
            await _excelHelper.SelectHeutigesDatumEntwicklungsBerichtAsync(UpdateMonatsrechnerCheckBox, group);

        if (success) return;
        UpdateMonatsrechnerCheckBox.IsChecked = false;
    }

    private void KidNameComboBox_Loaded(object sender, RoutedEventArgs e)
    {
        _autoCompleteHelper.KidNameComboBox_Loaded();
    }

    private void OnGroupSelected(object sender, SelectionChangedEventArgs e)
    {
        _autoCompleteHelper.OnGroupSelected(e);
    }

    private void KidNameComboBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        _autoCompleteHelper.OnKidNameComboBoxPreviewTextInput(e);
    }

    private void KidNameComboBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (sender is null) throw new ArgumentNullException(nameof(sender));
        _autoCompleteHelper.OnKidNameComboBoxPreviewKeyDown(e);
    }

    private void KidNameComboBox_LostFocus(object sender, RoutedEventArgs e)
    {
        KidNameComboBox.IsDropDownOpen = false;
    }

    private async void OnGenerateButtonClickedAsync(object sender, RoutedEventArgs e)
    {
        if (!await IsValidInputAsync()) return;

        await PerformFileOperationsAsync();
    }

    private void OnAboutClicked(object sender, RoutedEventArgs e)
    {
        var version = Assembly.GetEntryAssembly()?.GetName().Version;
        var versionString = version != null ? $"{version.Major}.{version.Minor}.{version.Build}" : "Unbekannt";

        var aboutWindow = new AboutWindow(versionString);
        aboutWindow.ShowDialog();
    }

    private async void OnProtokollbogenAutoCheckboxChanged(object sender, RoutedEventArgs e)
    {
        if (_isHandlingCheckboxEvent) return;

        _isHandlingCheckboxEvent = true;

        switch (e.RoutedEvent.Name)
        {
            case "Checked":
                await HandleProtokollbogenAutoCheckboxAsync(true);
                break;
            case "Unchecked":
                await HandleProtokollbogenAutoCheckboxAsync(false);
                break;
        }

        _isHandlingCheckboxEvent = false;
        e.Handled = true;
    }

    private async Task HandleProtokollbogenAutoCheckboxAsync(bool isChecked)
    {
        if (!isChecked) return;
        var group = GroupDropdown.Text;
        var kidName = KidNameComboBox.Text;

        var nameParts = kidName.Split(' ');

        if (nameParts.Length > 0)
        {
            var kidFirstName = nameParts[0].Trim();
            var kidLastName = "";

            for (var i = 1; i < nameParts.Length; i++) kidLastName += nameParts[i].Trim() + " ";

            kidLastName = kidLastName.Trim();

            var (months, error, _, _) =
                await _excelHelper.ExtractFromExcelAsync(group, kidLastName, kidFirstName);

            switch (error)
            {
                case "HomeFolderNotSet":
                    ShowMessage("Bitte setzen Sie zuerst den Heimordner.", MessageType.Error);
                    ProtokollbogenAutoCheckbox.IsChecked = false;
                    return;
                case "FileNotFound":
                    ShowMessage(
                        "Das erforderliche Excel-Dokument konnte nicht gefunden werden. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut.",
                        MessageType.Error);
                    ProtokollbogenAutoCheckbox.IsChecked = false;
                    return;
            }

            if (months.HasValue) return;
            ShowMessage("Das Alter des Kindes konnte nicht aus Excel extrahiert werden.",
                MessageType.Error);
            ProtokollbogenAutoCheckbox.IsChecked = false;
        }
        else
        {
            ShowMessage("Ungültiger Name. Bitte überprüfen Sie die Daten.", MessageType.Error);
            ProtokollbogenAutoCheckbox.IsChecked = false;
        }
    }

    private async Task<bool> IsValidInputAsync()
    {
        if (!IsHomeFolderSelected() || !AreAllRequiredFieldsFilled())
            return false;

        if (_fileManager == null)
        {
            if (HomeFolder != null)
            {
                _fileManager = new FileManagerHelper(HomeFolder);
            }
            else
            {
                ShowMessage("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", MessageType.Error);
                return false;
            }
        }

        if (HomeFolder == null)
            ShowMessage("Bitte wählen Sie zunächst das Hauptverzeichnis aus.",
                MessageType.Error);

        var kidName = KidNameComboBox.Text;
        if (HomeFolder != null)
        {
            var validatedKidName = await ValidationHelper.ValidateKidNameAsync(kidName, HomeFolder, GroupDropdown.Text);
            if (string.IsNullOrEmpty(validatedKidName))
            {
                ShowMessage("Ungültiger Kinder-Name", MessageType.Error);
                return false;
            }
        }

        var reportYearText = ReportYearTextbox.Text;
        try
        {
            var parsedYear = ValidationHelper.ValidateReportYearFromTextbox(reportYearText);
            if (!parsedYear.HasValue)
            {
                ShowMessage("Bitte geben Sie ein gültiges Jahr für den Bericht an.", MessageType.Error);
                return false;
            }
        }
        catch (Exception ex)
        {
            ShowMessage(
                $"Beim Verarbeiten der Excel-Datei ist ein unerwarteter Fehler aufgetreten: {ex.Message}",
                MessageType.Error);
            return false;
        }

        return true;
    }

    private static string? ExtractProtocolNumberFromData((string directoryPath, string fileName)? protokollbogenData)
    {
        if (!protokollbogenData.HasValue) return null;

        var fileName = protokollbogenData.Value.fileName + ".pdf";
        var match = ProtokollbogenFileNameRegex().Match(fileName);

        if (match.Success) return match.Groups[1].Value + "_Monate";
        ShowMessage("Fehler beim Extrahieren der Protokollnummer.", MessageType.Error);
        return null;
    }

    private bool ValidateHomeFolder()
    {
        if (HomeFolder != null) return true;
        ShowMessage("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", MessageType.Error);
        return false;
    }

    private async Task<string?> ValidateKidNameAsync()
    {
        var validatedKidName =
            await ValidationHelper.ValidateKidNameAsync(KidNameComboBox.Text, HomeFolder!, GroupDropdown.Text);
        if (validatedKidName == null) ShowMessage("Ungültiger Kinder-Name.", MessageType.Error);
        return validatedKidName;
    }

    private int? ValidateReportYear()
    {
        var reportYearNullable = ValidationHelper.ValidateReportYearFromTextbox(ReportYearTextbox.Text);
        if (!reportYearNullable.HasValue) ShowMessage("Ungültiges Jahr.", MessageType.Error);
        return reportYearNullable;
    }

    private static async Task FillPdfDocumentsAsync(string? protokollbogenPath,
        string? allgemeinEntwicklungsberichtPath,
        string? protokollElterngespraechPath, string? vorschuleEntwicklungsberichtPath, string? krippeUebergangsberichtPath, string kidName,
        double? months, string group, string parsedBirthDate, string? genderValue)
    {
        if (!string.IsNullOrEmpty(protokollbogenPath))
            await FillPdfAsync(protokollbogenPath, kidName, months ?? 0, group, PdfType.Protokollbogen,
                parsedBirthDate, genderValue);

        if (!string.IsNullOrEmpty(allgemeinEntwicklungsberichtPath))
            await FillPdfAsync(allgemeinEntwicklungsberichtPath, kidName, months ?? 0, group,
                PdfType.AllgemeinEntwicklungsbericht, parsedBirthDate, genderValue);

        if (!string.IsNullOrEmpty(protokollElterngespraechPath))
            await FillPdfAsync(protokollElterngespraechPath, kidName, months ?? 0, group,
                PdfType.ProtokollElterngespraech, parsedBirthDate, genderValue);

        if (!string.IsNullOrEmpty(vorschuleEntwicklungsberichtPath))
            await FillPdfAsync(vorschuleEntwicklungsberichtPath, kidName, months ?? 0, group,
                PdfType.VorschuleEntwicklungsbericht, parsedBirthDate, genderValue);

        if (!string.IsNullOrEmpty(krippeUebergangsberichtPath))
            await FillPdfAsync(krippeUebergangsberichtPath, kidName, months ?? 0, group,
                PdfType.KrippeUebergangsbericht, parsedBirthDate, genderValue);
    }

    private async Task CopyRequiredFilesAsync((string directoryPath, string fileName)? protokollbogenData,
        string sourceFolderPath,
        string targetFolderPath, string homeFolder, bool isAllgemeinerChecked, bool isVorschuleChecked,
        bool isProtokollbogenChecked, bool isKrippeUebergangsChecked)
    {
        if (_fileManager == null) throw new InvalidOperationException("_fileManager has not been initialized.");

        if (protokollbogenData.HasValue && !string.IsNullOrEmpty(sourceFolderPath) &&
            !string.IsNullOrEmpty(protokollbogenData.Value.fileName))
            if (isProtokollbogenChecked)
                await FileManagerHelper.CopyFilesFromSourceToTargetAsync(
                    Path.Combine(sourceFolderPath, protokollbogenData.Value.fileName + ".pdf"), targetFolderPath,
                    protokollbogenData.Value.fileName + ".pdf");

        var allgemeinerFilePath = Path.Combine(homeFolder, "Entwicklungsboegen", "Allgemeiner-Entwicklungsbericht.pdf");

        if (isAllgemeinerChecked && File.Exists(allgemeinerFilePath))
            await FileManagerHelper.CopyFilesFromSourceToTargetAsync(allgemeinerFilePath, targetFolderPath,
                Path.GetFileName(allgemeinerFilePath));
        else if (!File.Exists(allgemeinerFilePath))
            LogMessage(
                $"File 'Allgemeiner-Entwicklungsbericht.pdf' not found at {allgemeinerFilePath}.", LogLevel.Warning);

        var vorschuleFilePath = Path.Combine(homeFolder, "Entwicklungsboegen", "Vorschule-Entwicklungsbericht.pdf");

        if (isVorschuleChecked && File.Exists(vorschuleFilePath))
            await FileManagerHelper.CopyFilesFromSourceToTargetAsync(vorschuleFilePath, targetFolderPath,
                Path.GetFileName(vorschuleFilePath));
        else if (!File.Exists(vorschuleFilePath))
            LogMessage($"Datei 'Vorschule-Entwicklungsbericht.pdf' wurde nicht in {vorschuleFilePath} gefunden.",
                LogLevel.Warning);

        var protokollElterngespraechPath =
            Path.Combine(homeFolder, "Entwicklungsboegen", "Protokoll-Elterngespraech.pdf");

        if (File.Exists(protokollElterngespraechPath))
            await FileManagerHelper.CopyFilesFromSourceToTargetAsync(protokollElterngespraechPath, targetFolderPath,
                Path.GetFileName(protokollElterngespraechPath));
        else
            LogMessage(
                $"Datei 'Protokoll-Elterngespraech.pdf' wurde nicht in {protokollElterngespraechPath} gefunden.",
                LogLevel.Warning);

        var krippeÜbergangsberichtPath = 
            Path.Combine(homeFolder, "Entwicklungsboegen", "Krippe-Uebergangsbericht.pdf");

        if (isKrippeUebergangsChecked && File.Exists(krippeÜbergangsberichtPath))
            await FileManagerHelper.CopyFilesFromSourceToTargetAsync(krippeÜbergangsberichtPath, targetFolderPath,
                Path.GetFileName(krippeÜbergangsberichtPath));
        else if (!File.Exists(krippeÜbergangsberichtPath))
            LogMessage($"Datei 'Krippe-Uebergangsbericht.pdf' wurde nicht in {krippeÜbergangsberichtPath} gefunden.",
                LogLevel.Warning);

    }

    private async Task PerformFileOperationsAsync()
    {
        if (_fileManager == null)
        {
            ShowMessage("Der Dateimanager ist nicht initialisiert.", MessageType.Error);
            return;
        }

        OperationState.OperationsSuccessful = true;
        var sourceFolderPath = string.Empty;
        (string directoryPath, string fileName)? protokollbogenData = null;

        var group = ConvertToTitleCase(GroupDropdown.Text);

        if (!ValidateHomeFolder()) return;

        var validatedKidName = await ValidateKidNameAsync();
        if (validatedKidName == null) return;

        var kidName = ConvertToTitleCase(validatedKidName);
        var reportMonth = ConvertToTitleCase(ReportMonthDropdown.Text);

        var reportYearNullable = ValidateReportYear();
        if (!reportYearNullable.HasValue) return;
        var reportYear = reportYearNullable.Value;

        var nameParts = kidName.Split(' ');
        var kidFirstName = nameParts[0];
        var kidLastName = nameParts[1];

        var (months, _, parsedBirthDate, genderValue) =
            await _excelHelper.ExtractFromExcelAsync(group, kidLastName, kidFirstName);

        if (months.HasValue)
        {
            var formattedMonthsAndDays = ValidationHelper.ConvertToDecimalFormat(months.Value);
            var protokollbogenResult = ValidationHelper.DetermineProtokollbogen(formattedMonthsAndDays);
            if (protokollbogenResult.HasValue)
            {
                protokollbogenData = protokollbogenResult;

                if (HomeFolder == null)
                {
                    ShowMessage("Das Hauptverzeichnis ist nicht festgelegt.", MessageType.Error);
                    return;
                }

                var cleanedHomeFolder = HomeFolder.TrimEnd('\\');
                var cleanedDirectoryPath = protokollbogenData.Value.directoryPath.TrimStart('\\');

                sourceFolderPath = Path.Combine(cleanedHomeFolder, cleanedDirectoryPath);
            }
        }
        else
        {
            ShowMessage("Fehler beim Extrahieren der Monate aus Excel.", MessageType.Error);
            return;
        }

        var targetFolderPath = _fileManager.GetTargetPath(group, kidName, reportYear.ToString(), reportMonth);

        var isAllgemeinerChecked = AllgemeinerEntwicklungsberichtCheckbox.IsChecked == true;
        var isVorschuleChecked = VorschulentwicklungsberichtCheckbox.IsChecked == true;
        var isProtokollbogenChecked = ProtokollbogenAutoCheckbox.IsChecked == true;
        var isKrippeUebergangsberichtChecked = KrippeUebergangsberichtCheckbox.IsChecked == true;


        await CopyRequiredFilesAsync(protokollbogenData, sourceFolderPath, targetFolderPath, HomeFolder!,
            isAllgemeinerChecked,
            isVorschuleChecked, isProtokollbogenChecked, isKrippeUebergangsberichtChecked);

        var numericProtokollNumber = ExtractProtocolNumberFromData(protokollbogenData) ?? string.Empty;

        var (renamedProtokollbogenPath, renamedAllgemeinEntwicklungsberichtPath, renamedProtokollElterngespraechPath,
            renamedVorschuleEntwicklungsberichtPath, renamedKrippeUebergangsberichtPath) = await FileManagerHelper.RenameFilesInTargetDirectoryAsync(
            targetFolderPath,
            kidName, reportMonth, reportYear.ToString(), isAllgemeinerChecked, isVorschuleChecked,
            isProtokollbogenChecked, isKrippeUebergangsberichtChecked, numericProtokollNumber);


        if (string.IsNullOrEmpty(parsedBirthDate))
        {
            LogAndShowMessage("Geburtsdatum konnte nicht extrahiert werden.",
                "Fehler beim Extrahieren des Geburtsdatums.");
            return;
        }

        await FillPdfDocumentsAsync(renamedProtokollbogenPath, renamedAllgemeinEntwicklungsberichtPath,
            renamedProtokollElterngespraechPath, renamedVorschuleEntwicklungsberichtPath, renamedKrippeUebergangsberichtPath,
            kidName, months, group, parsedBirthDate, genderValue);

        if (OperationState.OperationsSuccessful)
            ShowMessage("Dateien erfolgreich kopiert und umbenannt.");
    }

    private bool IsHomeFolderSelected()
    {
        if (!string.IsNullOrEmpty(HomeFolder))
            return true;

        SelectHomeFolder();
        return !string.IsNullOrEmpty(HomeFolder);
    }

    private bool AreAllRequiredFieldsFilled()
    {
        var selectedGroup = GroupDropdown.Text;
        var childName = KidNameComboBox.Text;
        var selectedReportMonth = ReportMonthDropdown.Text;
        var selectedReportYear = ReportYearTextbox.Text;

        if (string.IsNullOrWhiteSpace(childName) || !childName.Contains(' '))
        {
            ShowMessage("Bitte geben Sie einen gültigen Namen mit Vor- und Nachnamen an.",
                MessageType.Error);
            return false;
        }

        if (!string.IsNullOrWhiteSpace(selectedGroup) && !string.IsNullOrWhiteSpace(selectedReportMonth) &&
            !string.IsNullOrWhiteSpace(selectedReportYear)) return true;
        ShowMessage("Bitte füllen Sie alle geforderten Felder aus.", MessageType.Error);
        return false;
    }

    private async void SelectHomeFolder()
    {
        var dialog = new VistaFolderBrowserDialog();
        {
            dialog.Description = "Wählen Sie das Hauptverzeichnis aus";
            dialog.UseDescriptionForTitle = true;

            var result = dialog.ShowDialog();
            if (!result.HasValue || !result.Value) return;
            HomeFolder = dialog.SelectedPath;
            InitializeFileManager();

            var settings = new Settings
            {
                HomeFolderPath = HomeFolder
            };
            await Settings.SaveSettingsAsync(settings);

            Dispatcher.Invoke(() => { ShowMessage($"Ausgewähltes Hauptverzeichnis: {HomeFolder}"); });
        }
    }

    private void MainWindow_Closed(object sender, EventArgs e)
    {
        Log.CloseAndFlush();
    }

    [GeneratedRegex(@"Kind_Protokollbogen_(\d+)_Monate\.pdf")]
    private static partial Regex ProtokollbogenFileNameRegex();

    public static class OperationState
    {
        public static bool OperationsSuccessful { get; set; } = true;
    }
}