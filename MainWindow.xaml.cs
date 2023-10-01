using MahApps.Metro.Controls;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static Automatisiertes_Kopieren.FileManager.StringUtilities;
using static Automatisiertes_Kopieren.FillPDF;
using static Automatisiertes_Kopieren.LoggingService;

namespace Automatisiertes_Kopieren
{
    public partial class MainWindow : MetroWindow
    {
        private readonly static LoggingService _loggingService = new LoggingService();
        private AutoComplete _autoComplete;
        private FileManager? _fileManager;
        private int? _selectedProtokollbogenMonth;
        private bool _isHandlingCheckboxEvent = false;
        private int _previousGroupSelectionIndex = 0;
        private List<string> _allKidNames = new List<string>();
        private ExcelService _excelService;

        public MainWindow()
        {
            _loggingService.InitializeLogger();
            InitializeComponent();
            _autoComplete = new AutoComplete(this);
            var settings = new AppSettings().LoadSettings();
            if (settings != null && !string.IsNullOrEmpty(settings.HomeFolderPath))
            {
                homeFolder = settings.HomeFolderPath;
            }
            else
            {
                SelectHomeFolder();
                if (string.IsNullOrEmpty(homeFolder))
                {
                    _loggingService.ShowMessage("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", MessageType.Error);
                    throw new InvalidOperationException("Home folder must be set.");
                }
            }
            _excelService = new ExcelService(homeFolder);

            protokollbogenAutoCheckbox.Checked += OnProtokollbogenAutoCheckboxChanged;
            protokollbogenAutoCheckbox.Unchecked += OnProtokollbogenAutoCheckboxChanged;
        }

        private string? _homeFolder;
        public string? homeFolder
        {
            get => _homeFolder;
            set
            {
                _homeFolder = value;
                if (groupDropdown.SelectedIndex == 0 && !string.IsNullOrEmpty(_homeFolder))
                {
                    var defaultKidNames = _autoComplete.GetKidNamesForGroup("Bären");
                    kidNameComboBox.ItemsSource = defaultKidNames;
                }
            }
        }

        private void OnSelectHomeFolderButtonClicked(object sender, RoutedEventArgs e)
        {
            SelectHomeFolder();
        }

        private void InitializeFileManager()
        {
            if (homeFolder != null)
            {
                _fileManager = new FileManager(homeFolder);
            }
            else
            {
                _loggingService.ShowMessage("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", MessageType.Error);
            }
        }

        private void OnSelectHeutigesDatumEntwicklungsBericht(object sender, RoutedEventArgs e)
        {
            _excelService.SelectHeutigesDatumEntwicklungsBericht(sender, e);
        }

        private void KidNameComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            _autoComplete.KidNameComboBox_Loaded(sender, e);
        }

        private void KidNameComboBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            _autoComplete.OnKidNameComboBoxPreviewTextInput(sender, e);
        }

        private void KidNameComboBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            _autoComplete.OnKidNameComboBoxPreviewKeyDown(sender, e);
        }

        private void OnProtokollbogenAutoCheckboxChanged(object sender, RoutedEventArgs e)
        {
            if (_isHandlingCheckboxEvent) return;

            _isHandlingCheckboxEvent = true;

            if (e.RoutedEvent.Name == "Checked")
            {
                HandleProtokollbogenAutoCheckbox(true);
            }
            else if (e.RoutedEvent.Name == "Unchecked")
            {
                HandleProtokollbogenAutoCheckbox(false);
            }

            _isHandlingCheckboxEvent = false;
            e.Handled = true;
        }

        private void HandleProtokollbogenAutoCheckbox(bool isChecked)
        {
            if (isChecked)
            {
                string group = groupDropdown.Text;
                string kidName = kidNameComboBox.Text;

                string[] nameParts = kidName.Split(' ');

                if (nameParts.Length > 0)
                {
                    string kidFirstName = nameParts[0].Trim();
                    string kidLastName = "";

                    for (int i = 1; i < nameParts.Length; i++)
                    {
                        kidLastName += nameParts[i].Trim() + " ";
                    }

                    kidLastName = kidLastName.Trim();
                    var result = _excelService.ExtractFromExcel(group, kidLastName, kidFirstName);

                    if (result.error == "HomeFolderNotSet")
                    {
                        _loggingService.ShowMessage("Bitte setzen Sie zuerst den Heimordner.", MessageType.Error);
                        protokollbogenAutoCheckbox.IsChecked = false;
                        return;
                    }
                    else if (result.error == "FileNotFound")
                    {
                        _loggingService.ShowMessage("Das erforderliche Excel-Dokument konnte nicht gefunden werden. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut.", MessageType.Error);
                        protokollbogenAutoCheckbox.IsChecked = false;
                        return;
                    }
                    else if (!result.months.HasValue)
                    {
                        _loggingService.ShowMessage("Das Alter des Kindes konnte nicht aus Excel extrahiert werden.", MessageType.Error);
                        protokollbogenAutoCheckbox.IsChecked = false;
                        return;
                    }
                    _selectedProtokollbogenMonth = (int)Math.Round(result.months.Value);
                }
                else
                {
                    _loggingService.ShowMessage("Ungültiger Name. Bitte überprüfen Sie die Daten.", MessageType.Error);
                    protokollbogenAutoCheckbox.IsChecked = false;
                }
            }
            else
            {
                _selectedProtokollbogenMonth = null;
            }
        }

        private void OnGenerateButtonClicked(object sender, RoutedEventArgs e)
        {
            if (!IsValidInput())
            {
                return;
            }

            PerformFileOperations();
        }

        private bool IsValidInput()
        {
            if (!IsHomeFolderSelected() || !AreAllRequiredFieldsFilled())
                return false;

            if (_fileManager == null)
            {
                if (homeFolder != null)
                {
                    _fileManager = new FileManager(homeFolder);
                }
                else
                {
                    _loggingService.ShowMessage("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", MessageType.Error);
                    return false;
                }
            }

            if (homeFolder == null)
            {
                _loggingService.ShowMessage("Das Hauptverzeichnis ist nicht festgelegt.", MessageType.Error);
                return false;
            }

            string kidName = kidNameComboBox.Text;
            string? validatedKidName = ValidationHelper.ValidateKidName(kidName, homeFolder, groupDropdown.Text);
            if (string.IsNullOrEmpty(validatedKidName))
            {
                _loggingService.ShowMessage("Ungültiger Kinder-Name", MessageType.Error);
                return false;
            }

            string reportYearText = reportYearTextbox.Text;
            try
            {
                int? parsedYear = ValidationHelper.ValidateReportYearFromTextbox(reportYearText);
                if (!parsedYear.HasValue)
                {
                    _loggingService.ShowMessage("Bitte geben Sie ein gültiges Jahr für den Bericht an.", MessageType.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                _loggingService.ShowMessage($"Beim Verarbeiten der Excel-Datei ist ein unerwarteter Fehler aufgetreten: {ex.Message}", MessageType.Error);
                return false;
            }

            return true;
        }

        private string? ExtractProtokollNumberFromData((string directoryPath, string fileName)? protokollbogenData)
        {
            if (!protokollbogenData.HasValue)
            {
                return null;
            }

            var fileName = protokollbogenData.Value.fileName + ".pdf";
            var match = Regex.Match(fileName, @"Kind_Protokollbogen_(\d+)_Monate\.pdf");

            if (!match.Success)
            {
                _loggingService.ShowMessage("Fehler beim Extrahieren der Protokollnummer.", MessageType.Error);
                return null;
            }

            return match.Groups[1].Value + "_Monate";
        }

        public static class OperationState
        {
            public static bool OperationsSuccessful { get; set; } = true;
        }

        private bool ValidateHomeFolder()
        {
            if (homeFolder == null)
            {
                _loggingService.ShowMessage("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", MessageType.Error);
                return false;
            }
            return true;
        }

        private string? ValidateKidName()
        {
            string? validatedKidName = ValidationHelper.ValidateKidName(kidNameComboBox.Text, homeFolder!, groupDropdown.Text);
            if (validatedKidName == null)
            {
                _loggingService.ShowMessage("Ungültiger Kinder-Name.", MessageType.Error);
            }
            return validatedKidName;
        }

        private int? ValidateReportYear()
        {
            int? reportYearNullable = ValidationHelper.ValidateReportYearFromTextbox(reportYearTextbox.Text);
            if (!reportYearNullable.HasValue)
            {
                _loggingService.ShowMessage("Ungültiges Jahr.", MessageType.Error);
            }
            return reportYearNullable;
        }

        private void FillPdfDocuments(string? protokollbogenPath, string? allgemeinEntwicklungsberichtPath, string? protokollElterngespraechFilePath, string? vorschulEntwicklungsberichtPath, string kidName, double? months, string group, string parsedBirthDate, string? genderValue)
        {
            var fillPdf = new FillPDF();

            if (!string.IsNullOrEmpty(protokollbogenPath))
            {
                fillPdf.FillPdf(protokollbogenPath, kidName, months.HasValue ? months.Value : 0, group, PdfType.Protokollbogen, parsedBirthDate, genderValue);
            }

            if (!string.IsNullOrEmpty(allgemeinEntwicklungsberichtPath))
            {
                fillPdf.FillPdf(allgemeinEntwicklungsberichtPath, kidName, months.HasValue ? months.Value : 0, group, PdfType.AllgemeinEntwicklungsbericht, parsedBirthDate, genderValue);
            }

            if (!string.IsNullOrEmpty(protokollElterngespraechFilePath))
            {
                fillPdf.FillPdf(protokollElterngespraechFilePath, kidName, months.HasValue ? months.Value : 0, group, PdfType.ProtokollElterngespraech, parsedBirthDate, genderValue);
            }
            if (!string.IsNullOrEmpty(vorschulEntwicklungsberichtPath))
            {
                fillPdf.FillPdf(vorschulEntwicklungsberichtPath, kidName, months.HasValue ? months.Value : 0, group, PdfType.VorschulEntwicklungsbericht, parsedBirthDate, genderValue);
            }
        }

        private void CopyRequiredFiles((string directoryPath, string fileName)? protokollbogenData, string sourceFolderPath, string targetFolderPath, string homeFolder, bool isAllgemeinerChecked, bool isVorschulChecked, bool isProtokollbogenChecked)
        {

            if (_fileManager == null)
            {
                throw new InvalidOperationException("_fileManager has not been initialized.");
            }

            if (protokollbogenData.HasValue && !string.IsNullOrEmpty(sourceFolderPath) && !string.IsNullOrEmpty(protokollbogenData.Value.fileName))
            {
                if (isProtokollbogenChecked)
                {
                    _fileManager.CopyFilesFromSourceToTarget(Path.Combine(sourceFolderPath, protokollbogenData.Value.fileName + ".pdf"), targetFolderPath, protokollbogenData.Value.fileName + ".pdf");
                }
            }

            string allgemeinerFilePath = Path.Combine(homeFolder, "Entwicklungsboegen", "Allgemeiner-Entwicklungsbericht.pdf");

            if (isAllgemeinerChecked && File.Exists(allgemeinerFilePath))
            {
                _fileManager.CopyFilesFromSourceToTarget(allgemeinerFilePath, targetFolderPath, Path.GetFileName(allgemeinerFilePath) ?? string.Empty);
            }
            else if (!File.Exists(allgemeinerFilePath))
            {
                _loggingService.LogMessage($"File 'Allgemeiner-Entwicklungsbericht.pdf' not found at {allgemeinerFilePath}.", LogLevel.Warning);
            }

            string vorschulFilePath = Path.Combine(homeFolder, "Entwicklungsboegen", "Vorschul-Entwicklungsbericht.pdf");

            if (isVorschulChecked && File.Exists(vorschulFilePath))
            {
                _fileManager.CopyFilesFromSourceToTarget(vorschulFilePath, targetFolderPath, Path.GetFileName(vorschulFilePath) ?? string.Empty);
            }
            else if (!File.Exists(vorschulFilePath))
            {
                _loggingService.LogMessage($"File 'Vorschul-Entwicklungsbericht.pdf' not found at {vorschulFilePath}.", LogLevel.Warning);
            }

            string protokollElterngespraechFilePath = Path.Combine(homeFolder, "Entwicklungsboegen", "Protokoll-Elterngespraech.pdf");

            if (File.Exists(protokollElterngespraechFilePath))
            {
                _fileManager.CopyFilesFromSourceToTarget(protokollElterngespraechFilePath, targetFolderPath, Path.GetFileName(protokollElterngespraechFilePath) ?? string.Empty);
            }
            else
            {
                _loggingService.LogMessage($"File 'Protokoll-Elterngespraech.pdf' not found at {protokollElterngespraechFilePath}.", LogLevel.Warning);
            }
        }

        private void PerformFileOperations()
        {
            if (_fileManager == null)
            {
                _loggingService.ShowMessage("Der Dateimanager ist nicht initialisiert.", MessageType.Error);
                return;
            }
            OperationState.OperationsSuccessful = true;
            string sourceFolderPath = string.Empty;
            (string directoryPath, string fileName)? protokollbogenData = null;
            string numericProtokollNumber = string.Empty;

            string group = ConvertToTitleCase(groupDropdown.Text);

            if (!ValidateHomeFolder()) return;

            string? validatedKidName = ValidateKidName();
            if (validatedKidName == null) return;

            string kidName = ConvertToTitleCase(validatedKidName);
            string reportMonth = ConvertToTitleCase(reportMonthDropdown.Text);

            int? reportYearNullable = ValidateReportYear();
            if (!reportYearNullable.HasValue) return;
            int reportYear = reportYearNullable.Value;

            var nameParts = kidName.Split(' ');
            string kidFirstName = nameParts[0];
            string kidLastName = nameParts[1];

            var (months, error, parsedBirthDate, genderValue) = _excelService.ExtractFromExcel(group, kidLastName, kidFirstName);

            if (months.HasValue)
            {
                double formattedMonthsAndDays = ValidationHelper.ConvertToDecimalFormat(months.Value);
                var protokollbogenResult = ValidationHelper.DetermineProtokollbogen(formattedMonthsAndDays);
                if (protokollbogenResult.HasValue)
                {
                    protokollbogenData = protokollbogenResult;

                    if (homeFolder == null)
                    {
                        _loggingService.ShowMessage("Das Hauptverzeichnis ist nicht festgelegt.", MessageType.Error);
                        return;
                    }

                    string cleanedHomeFolder = homeFolder.TrimEnd('\\');
                    string cleanedDirectoryPath = protokollbogenData.Value.directoryPath.TrimStart('\\');

                    sourceFolderPath = Path.Combine(cleanedHomeFolder, cleanedDirectoryPath);
                }
            }
            else
            {
                _loggingService.ShowMessage("Fehler beim Extrahieren der Monate aus Excel.", MessageType.Error);
                return;
            }

            string targetFolderPath = _fileManager.GetTargetPath(group, kidName, reportYear.ToString(), reportMonth);

            bool isAllgemeinerChecked = allgemeinerEntwicklungsberichtCheckbox.IsChecked == true;
            bool isVorschulChecked = vorschulentwicklungsberichtCheckbox.IsChecked == true;
            bool isProtokollbogenChecked = protokollbogenAutoCheckbox.IsChecked == true;

            CopyRequiredFiles(protokollbogenData, sourceFolderPath, targetFolderPath, homeFolder!, isAllgemeinerChecked, isVorschulChecked, isProtokollbogenChecked);

            numericProtokollNumber = ExtractProtokollNumberFromData(protokollbogenData) ?? string.Empty;

            var (renamedProtokollbogenPath, renamedAllgemeinEntwicklungsberichtPath, renamedProtokollElterngespraechPath, renamedVorschulEntwicklungsberichtPath) = _fileManager.RenameFilesInTargetDirectory(targetFolderPath, kidName, reportMonth, reportYear.ToString(), isAllgemeinerChecked, isVorschulChecked, isProtokollbogenChecked, numericProtokollNumber);

            if (string.IsNullOrEmpty(parsedBirthDate))
            {
                _loggingService.LogAndShowMessage("Geburtsdatum konnte nicht extrahiert werden.", "Error extracting birth date.");
                return;
            }
            FillPdfDocuments(renamedProtokollbogenPath, renamedAllgemeinEntwicklungsberichtPath, renamedProtokollElterngespraechPath, renamedVorschulEntwicklungsberichtPath, kidName, months, group, parsedBirthDate, genderValue);
            if (OperationState.OperationsSuccessful)
            {
                _loggingService.ShowMessage("Dateien erfolgreich kopiert und umbenannt.", MessageType.Info);
            }
        }

        private bool IsHomeFolderSelected()
        {
            if (!string.IsNullOrEmpty(homeFolder))
                return true;

            SelectHomeFolder();
            return !string.IsNullOrEmpty(homeFolder);
        }

        private bool AreAllRequiredFieldsFilled()
        {
            string selectedGroup = groupDropdown.Text;
            string childName = kidNameComboBox.Text;
            string selectedReportMonth = reportMonthDropdown.Text;
            string selectedReportYear = reportYearTextbox.Text;

            if (string.IsNullOrWhiteSpace(childName) || !childName.Contains(" "))
            {
                _loggingService.ShowMessage("Bitte geben Sie einen gültigen Namen mit Vor- und Nachnamen an.", MessageType.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(selectedGroup) || string.IsNullOrWhiteSpace(selectedReportMonth) || string.IsNullOrWhiteSpace(selectedReportYear))
            {
                _loggingService.ShowMessage("Bitte füllen Sie alle geforderten Felder aus.", MessageType.Error);
                return false;
            }

            return true;
        }

        private void GroupDropdown_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? selectedGroup = groupDropdown.SelectedItem?.ToString();

            if (!string.IsNullOrEmpty(selectedGroup))
            {
                List<string> kidNames = _autoComplete.GetKidNamesForGroup(selectedGroup);

                kidNameComboBox.ItemsSource = kidNames;
            }
        }

        public void OnGroupSelected(object sender, SelectionChangedEventArgs e)
        {
            if (kidNameComboBox == null)
            {
                return;
            }

            if (string.IsNullOrEmpty(homeFolder))
            {
                MessageBoxResult result = _loggingService.ShowMessage("Möchten Sie das Hauptverzeichnis ändern?", MessageType.Info, "Hauptverzeichnis nicht festgelegt", MessageBoxButton.YesNo);

                if (result == MessageBoxResult.Yes)
                {
                    using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
                    {
                        System.Windows.Forms.DialogResult dialogResult = dialog.ShowDialog();
                        if (dialogResult == System.Windows.Forms.DialogResult.OK)
                        {
                            homeFolder = dialog.SelectedPath;
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                else
                {
                    return;
                }
            }

            _previousGroupSelectionIndex = groupDropdown.SelectedIndex;

            if (e.AddedItems.Count > 0 && e.AddedItems[0] is ComboBoxItem comboBoxItem && comboBoxItem.Content is string selectedGroup && !string.IsNullOrEmpty(selectedGroup))
            {
                var kidNames = _autoComplete.GetKidNamesForGroup(selectedGroup);
                _allKidNames = _autoComplete.GetKidNamesForGroup(selectedGroup);
                kidNameComboBox.ItemsSource = _allKidNames;
            }
            else if (e.AddedItems.Count > 0)
            {
                _loggingService.LogMessage($"e.AddedItems[0] type: {e.AddedItems[0]?.GetType().Name ?? "null"}, value: {e.AddedItems[0]}", LogLevel.Warning);
                _loggingService.LogMessage("Selected group is empty or not a valid ComboBoxItem.", LogLevel.Warning);
            }
            else
            {
                _loggingService.LogMessage("No group selected.", LogLevel.Warning);
            }
        }

        private void SelectHomeFolder()
        {
            var dialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            {
                dialog.Description = "Wählen Sie das Hauptverzeichnis aus";
                dialog.UseDescriptionForTitle = true;

                var result = dialog.ShowDialog();
                if (result.HasValue && result.Value)
                {
                    homeFolder = dialog.SelectedPath;
                    InitializeFileManager();

                    var settings = new AppSettings
                    {
                        HomeFolderPath = homeFolder
                    };
                    settings.SaveSettings(settings);

                    _loggingService.ShowMessage($"Ausgewähltes Hauptverzeichnis: {homeFolder}", MessageType.Info);
                }
            }
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            Log.CloseAndFlush();
        }
    }
}
