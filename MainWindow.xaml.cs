using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Ookii.Dialogs.Wpf;
using Serilog;
using static Automatisiertes_Kopieren.FileManager.StringUtilities;
using static Automatisiertes_Kopieren.PdfHelper;
using static Automatisiertes_Kopieren.LoggingHelper;

namespace Automatisiertes_Kopieren;

public partial class MainWindow
{
    private readonly AutoCompleteHelper _autoCompleteHelper;
    private readonly ExcelHelper _excelHelper;
    private FileManager? _fileManager;

    private string? _homeFolder;
    private bool _isHandlingCheckboxEvent;

    public MainWindow()
    {
        InitializeLogger();
        _autoCompleteHelper = new AutoCompleteHelper(this);
        InitializeComponent();
        var settings = AppSettings.LoadSettings();
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
                throw new InvalidOperationException("Home folder must be set.");
            }
        }

        _excelHelper = new ExcelHelper(HomeFolder);

        ProtokollbogenAutoCheckbox.Checked += OnProtokollbogenAutoCheckboxChanged;
        ProtokollbogenAutoCheckbox.Unchecked += OnProtokollbogenAutoCheckboxChanged;
    }

    public string? HomeFolder
    {
        get => _homeFolder;
        private set
        {
            _homeFolder = value;
            if (GroupDropdown.SelectedIndex != 0 || string.IsNullOrEmpty(_homeFolder)) return;
            var defaultKidNames = _autoCompleteHelper.GetKidNamesForGroup("Bären");
            KidNameComboBox.ItemsSource = defaultKidNames;
        }
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
            _fileManager = new FileManager(HomeFolder);
        else
            ShowMessage("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", MessageType.Error);
    }

    private void OnSelectHeutigesDatumEntwicklungsBericht(object sender, RoutedEventArgs e)
    {
        _excelHelper.SelectHeutigesDatumEntwicklungsBericht(sender);
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

    private void KidNameComboBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (sender is null) throw new ArgumentNullException(nameof(sender));
        var argumentOutOfRangeException = new ArgumentOutOfRangeException();
        _autoCompleteHelper.OnKidNameComboBoxPreviewKeyDown(e, argumentOutOfRangeException);
    }

    private void OnProtokollbogenAutoCheckboxChanged(object sender, RoutedEventArgs e)
    {
        if (_isHandlingCheckboxEvent) return;

        _isHandlingCheckboxEvent = true;

        switch (e.RoutedEvent.Name)
        {
            case "Checked":
                HandleProtokollbogenAutoCheckbox(true);
                break;
            case "Unchecked":
                HandleProtokollbogenAutoCheckbox(false);
                break;
        }

        _isHandlingCheckboxEvent = false;
        e.Handled = true;
    }

    private void HandleProtokollbogenAutoCheckbox(bool isChecked)
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
                _excelHelper.ExtractFromExcel(group, kidLastName, kidFirstName);

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

    private void OnGenerateButtonClicked(object sender, RoutedEventArgs e)
    {
        if (!IsValidInput()) return;

        PerformFileOperations();
    }

    private bool IsValidInput()
    {
        if (!IsHomeFolderSelected() || !AreAllRequiredFieldsFilled())
            return false;

        if (_fileManager == null)
        {
            if (HomeFolder != null)
            {
                _fileManager = new FileManager(HomeFolder);
            }
            else
            {
                ShowMessage("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", MessageType.Error);
                return false;
            }
        }

        if (HomeFolder == null)
        {
            ShowMessage("Das Hauptverzeichnis ist nicht festgelegt.", MessageType.Error);
            return false;
        }

        var kidName = KidNameComboBox.Text;
        var validatedKidName = ValidationHelper.ValidateKidName(kidName, HomeFolder, GroupDropdown.Text);
        if (string.IsNullOrEmpty(validatedKidName))
        {
            ShowMessage("Ungültiger Kinder-Name", MessageType.Error);
            return false;
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

    private static string? ExtractProtokollNumberFromData((string directoryPath, string fileName)? protokollbogenData)
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

    private string? ValidateKidName()
    {
        var validatedKidName = ValidationHelper.ValidateKidName(KidNameComboBox.Text, HomeFolder!, GroupDropdown.Text);
        if (validatedKidName == null) ShowMessage("Ungültiger Kinder-Name.", MessageType.Error);
        return validatedKidName;
    }

    private int? ValidateReportYear()
    {
        var reportYearNullable = ValidationHelper.ValidateReportYearFromTextbox(ReportYearTextbox.Text);
        if (!reportYearNullable.HasValue) ShowMessage("Ungültiges Jahr.", MessageType.Error);
        return reportYearNullable;
    }

    private static void FillPdfDocuments(string? protokollbogenPath, string? allgemeinEntwicklungsberichtPath,
        string? protokollElterngespraechFilePath, string? vorschuleEntwicklungsberichtPath, string kidName,
        double? months, string group, string parsedBirthDate, string? genderValue)
    {
        if (!string.IsNullOrEmpty(protokollbogenPath))
            FillPdf(protokollbogenPath, kidName, months ?? 0, group, PdfType.Protokollbogen,
                parsedBirthDate, genderValue);

        if (!string.IsNullOrEmpty(allgemeinEntwicklungsberichtPath))
            FillPdf(allgemeinEntwicklungsberichtPath, kidName, months ?? 0, group,
                PdfType.AllgemeinEntwicklungsbericht, parsedBirthDate, genderValue);

        if (!string.IsNullOrEmpty(protokollElterngespraechFilePath))
            FillPdf(protokollElterngespraechFilePath, kidName, months ?? 0, group,
                PdfType.ProtokollElterngespraech, parsedBirthDate, genderValue);
        if (!string.IsNullOrEmpty(vorschuleEntwicklungsberichtPath))
            FillPdf(vorschuleEntwicklungsberichtPath, kidName, months ?? 0, group,
                PdfType.VorschuleEntwicklungsbericht, parsedBirthDate, genderValue);
    }

    private void CopyRequiredFiles((string directoryPath, string fileName)? protokollbogenData, string sourceFolderPath,
        string targetFolderPath, string homeFolder, bool isAllgemeinerChecked, bool isVorschuleChecked,
        bool isProtokollbogenChecked)
    {
        if (_fileManager == null) throw new InvalidOperationException("_fileManager has not been initialized.");

        if (protokollbogenData.HasValue && !string.IsNullOrEmpty(sourceFolderPath) &&
            !string.IsNullOrEmpty(protokollbogenData.Value.fileName))
            if (isProtokollbogenChecked)
                FileManager.CopyFilesFromSourceToTarget(
                    Path.Combine(sourceFolderPath, protokollbogenData.Value.fileName + ".pdf"), targetFolderPath,
                    protokollbogenData.Value.fileName + ".pdf");

        var allgemeinerFilePath = Path.Combine(homeFolder, "Entwicklungsboegen", "Allgemeiner-Entwicklungsbericht.pdf");

        if (isAllgemeinerChecked && File.Exists(allgemeinerFilePath))
            FileManager.CopyFilesFromSourceToTarget(allgemeinerFilePath, targetFolderPath,
                Path.GetFileName(allgemeinerFilePath));
        else if (!File.Exists(allgemeinerFilePath))
            LogMessage(
                $"File 'Allgemeiner-Entwicklungsbericht.pdf' not found at {allgemeinerFilePath}.", LogLevel.Warning);

        var vorschuleFilePath = Path.Combine(homeFolder, "Entwicklungsboegen", "Vorschule-Entwicklungsbericht.pdf");

        if (isVorschuleChecked && File.Exists(vorschuleFilePath))
            FileManager.CopyFilesFromSourceToTarget(vorschuleFilePath, targetFolderPath,
                Path.GetFileName(vorschuleFilePath));
        else if (!File.Exists(vorschuleFilePath))
            LogMessage($"File 'Vorschule-Entwicklungsbericht.pdf' not found at {vorschuleFilePath}.",
                LogLevel.Warning);

        var protokollElterngespraechFilePath =
            Path.Combine(homeFolder, "Entwicklungsboegen", "Protokoll-Elterngespraech.pdf");

        if (File.Exists(protokollElterngespraechFilePath))
            FileManager.CopyFilesFromSourceToTarget(protokollElterngespraechFilePath, targetFolderPath,
                Path.GetFileName(protokollElterngespraechFilePath));
        else
            LogMessage(
                $"File 'Protokoll-Elterngespraech.pdf' not found at {protokollElterngespraechFilePath}.",
                LogLevel.Warning);
    }

    private void PerformFileOperations()
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

        var validatedKidName = ValidateKidName();
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
            _excelHelper.ExtractFromExcel(group, kidLastName, kidFirstName);

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

        CopyRequiredFiles(protokollbogenData, sourceFolderPath, targetFolderPath, HomeFolder!, isAllgemeinerChecked,
            isVorschuleChecked, isProtokollbogenChecked);

        var numericProtokollNumber = ExtractProtokollNumberFromData(protokollbogenData) ?? string.Empty;

        var (renamedProtokollbogenPath, renamedAllgemeinEntwicklungsberichtPath, renamedProtokollElterngespraechPath,
            renamedVorschuleEntwicklungsberichtPath) = FileManager.RenameFilesInTargetDirectory(targetFolderPath,
            kidName, reportMonth, reportYear.ToString(), isAllgemeinerChecked, isVorschuleChecked,
            isProtokollbogenChecked, numericProtokollNumber);

        if (string.IsNullOrEmpty(parsedBirthDate))
        {
            LogAndShowMessage("Geburtsdatum konnte nicht extrahiert werden.",
                "Error extracting birth date.");
            return;
        }

        FillPdfDocuments(renamedProtokollbogenPath, renamedAllgemeinEntwicklungsberichtPath,
            renamedProtokollElterngespraechPath, renamedVorschuleEntwicklungsberichtPath, kidName, months, group,
            parsedBirthDate, genderValue);
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

    private void SelectHomeFolder()
    {
        var dialog = new VistaFolderBrowserDialog();
        {
            dialog.Description = "Wählen Sie das Hauptverzeichnis aus";
            dialog.UseDescriptionForTitle = true;

            var result = dialog.ShowDialog();
            if (!result.HasValue || !result.Value) return;
            HomeFolder = dialog.SelectedPath;
            InitializeFileManager();

            var settings = new AppSettings
            {
                HomeFolderPath = HomeFolder
            };
            AppSettings.SaveSettings(settings);

            ShowMessage($"Ausgewähltes Hauptverzeichnis: {HomeFolder}");
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