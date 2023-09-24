using ClosedXML.Excel;
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
using MessageBox = System.Windows.MessageBox;

namespace Automatisiertes_Kopieren
{
    public partial class MainWindow : MetroWindow
    {

        private FileManager? _fileManager;
        private AutoComplete _autoComplete;
        private int? _selectedProtokollbogenMonth;
        private bool _isHandlingCheckboxEvent = false;
        private bool _isWindowLoaded = false;
        private int _previousGroupSelectionIndex = 0;
        private List<string> _allKidNames = new List<string>();
        private ValidationHelper _validator;
        private readonly LoggingService _loggingService;
        private void InitializeFileManager()
        {
            if (HomeFolder != null && ValidationHelper.IsValidDirectoryPath(HomeFolder))
            {
                _fileManager = new FileManager(HomeFolder, this);
            }
            else
            {
                _loggingService.ShowError("Bitte wählen Sie zunächst das Hauptverzeichnis aus.");
            }
        }
        public MainWindow()
        {
            InitializeComponent();

            _autoComplete = new AutoComplete(this);
            _validator = new ValidationHelper(this);
            _loggingService = new LoggingService(this);

            string logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Automatisiertes_Kopieren", "logs");
            string logFilePath = Path.Combine(logDirectory, "log-.txt");

            Log.Logger = new LoggerConfiguration()
                .WriteTo.File(logFilePath, rollingInterval: RollingInterval.Day, retainedFileCountLimit: 10)
                .CreateLogger();

            var settings = new AppSettings().LoadSettings();
            if (settings != null && !string.IsNullOrEmpty(settings.HomeFolderPath))
            {
                HomeFolder = settings.HomeFolderPath;
            }
            else
            {
                SelectHomeFolder();
            }

            InitializeFileManager();

            this.Loaded += MainWindow_Loaded;

            protokollbogenAutoCheckbox.Checked += OnProtokollbogenAutoCheckboxChanged;
            protokollbogenAutoCheckbox.Unchecked += OnProtokollbogenAutoCheckboxChanged;
        }


        private string? _homeFolder;
        public string? HomeFolder
        {
            get => _homeFolder;
            set
            {
                _homeFolder = value;
                if (groupDropdown != null && groupDropdown.SelectedIndex == 0 && !string.IsNullOrEmpty(_homeFolder))
                {
                    var defaultKidNames = _autoComplete.GetKidNamesForGroup("Bären");
                    kidNameComboBox.ItemsSource = defaultKidNames;
                }
            }
        }

        private (double? months, string? error) ExtractMonthsFromExcel(string group, string lastName, string firstName)
        {
            if (string.IsNullOrEmpty(HomeFolder))
            {
                return (null, "HomeFolderNotSet");
            }
            string convertedGroupName = ConvertSpecialCharacters(group, ConversionType.Umlaute);
            string shortGroupName = convertedGroupName.Split(' ')[0];
            string filePath = $@"{HomeFolder}\Entwicklungsberichte\{convertedGroupName} Entwicklungsberichte\Monatsrechner-Kinder-Zielsetzung-{shortGroupName}.xlsm";

            if (!ValidationHelper.IsValidFilePath(filePath))
            {
                Log.Error($"Verzeichnis existiert nicht: {filePath}");
                return (null, "InvalidPath");
            }

            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet("Monatsrechner");

                    for (int row = 7; row <= 31; row++)
                    {
                        var lastNameCell = worksheet.Cell(row, 3).Value.ToString().Trim();
                        var firstNameCell = worksheet.Cell(row, 4).Value.ToString().Trim();

                        if (lastNameCell != lastName || firstNameCell != firstName)
                        {
                            continue;
                        }

                        var monthsValueRaw = worksheet.Cell(row, 6).Value.ToString();

                        if (double.TryParse(monthsValueRaw.Replace(",", "."), out double parsedValue))
                        {
                            return (parsedValue, null);
                        }
                    }
                }
            }
            catch (FileNotFoundException)
            {
                _loggingService.LogAndShowError($"Die Datei {filePath} wurde nicht gefunden.", "Das erforderliche Excel-Dokument konnte nicht gefunden werden. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut.");
            }
            catch (Exception ex)
            {
                _loggingService.LogAndShowError($"Beim Verarbeiten der Excel-Datei ist ein unerwarteter Fehler aufgetreten. {ex.Message}", "Ein unerwarteter Fehler ist aufgetreten. Bitte versuchen Sie es erneut.");
                return (null, "UnexpectedError");
            }

            _loggingService.LogAndShowError($"Es konnte kein gültiger Monatswert für {firstName} {lastName} extrahiert werden.", "Ein Fehler ist beim Extrahieren des Monatswertes aufgetreten.");
            return (null, "ExtractionError");
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _isWindowLoaded = true;
        }

        private void OnProtokollbogenAutoCheckboxChanged(object sender, RoutedEventArgs e)
        {
            if (_isHandlingCheckboxEvent) return;

            _isHandlingCheckboxEvent = true;

            if (e.RoutedEvent.Name == "Checked")
            {
                HandleCheckboxChecked();
            }
            else if (e.RoutedEvent.Name == "Unchecked")
            {
                HandleCheckboxUnchecked();
            }

            _isHandlingCheckboxEvent = false;
            e.Handled = true;
        }

        private void HandleCheckboxChecked()
        {
            string group = groupDropdown.Text;
            string kidName = kidNameComboBox.Text;
            var nameParts = kidName.Split(' ');
            string kidFirstName = nameParts[0];
            string kidLastName = nameParts.Length > 1 ? nameParts[1] : "";

            var result = ExtractMonthsFromExcel(group, kidLastName, kidFirstName);
            if (result.error == "HomeFolderNotSet")
            {
                _loggingService.ShowError("Bitte setzen Sie zuerst den Heimordner.");
                protokollbogenAutoCheckbox.IsChecked = false;
                return;
            }
            else if (result.error == "FileNotFound")
            {
                _loggingService.ShowError("Das erforderliche Excel-Dokument konnte nicht gefunden werden. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut.");
                protokollbogenAutoCheckbox.IsChecked = false;
                return;
            }
            else if (!result.months.HasValue)
            {
                _loggingService.ShowError("Das Alter des Kindes konnte nicht aus Excel extrahiert werden.");
                protokollbogenAutoCheckbox.IsChecked = false;
                return;
            }
            _selectedProtokollbogenMonth = (int)Math.Round(result.months.Value);
        }

        private void HandleCheckboxUnchecked()
        {
            _selectedProtokollbogenMonth = null;
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
            return IsHomeFolderSet() && AreAllRequiredFieldsFilled() && IsKidNameValid() && IsReportYearValid();
        }

        private bool IsHomeFolderSet()
        {
            if (string.IsNullOrEmpty(HomeFolder))
            {
                _loggingService.LogAndShowError("HomeFolder is not set.", "Bitte wählen Sie zunächst das Hauptverzeichnis aus.");
                return false;
            }
            return true;
        }

        private bool IsKidNameValid()
        {
            if (HomeFolder == null)
            {
                _loggingService.LogAndShowError("HomeFolder is null during kid name validation.", "Bitte wählen Sie zunächst das Hauptverzeichnis aus.");
                return false;
            }

            string kidName = kidNameComboBox.Text;
            string? validatedKidName = _validator.IsKidNameValid();

            if (string.IsNullOrEmpty(validatedKidName))
            {
                _loggingService.LogAndShowError($"Kid name validation failed for name: {kidName}", "Ungültiger Kindername. Bitte überprüfen Sie den eingegebenen Namen.");
                return false;
            }

            return true;
        }

        private bool IsReportYearValid()
        {
            string reportYearText = reportYearTextbox.Text;
            int? parsedYear = _validator.IsReportYearValid();

            if (!parsedYear.HasValue)
            {
                _loggingService.LogAndShowError($"Report year validation failed for year: {reportYearText}", "Ungültiges Jahr.");
                return false;
            }
            return true;
        }

        private static string ExtractProtokollNumber(string fileName)
        {
            var match = Regex.Match(fileName, @"Kind_Protokollbogen_(\d+)_Monate\.pdf");
            return match.Success ? match.Groups[1].Value + "_Monate" : string.Empty;
        }

        private void PerformFileOperations()
        {
            if (!InitializeAndValidate(out string kidName, out int reportYear))
                return;

            if (!PrepareFilePaths(kidName, reportYear, out string sourceFolderPath, out string targetFolderPath, out string numericProtokollNumber))
                return;

            ExecuteFileOperations(kidName, reportYear, sourceFolderPath, targetFolderPath, numericProtokollNumber);
        }

        private bool InitializeAndValidate(out string kidName, out int reportYear)
        {
            kidName = string.Empty;
            reportYear = 0;

            if (HomeFolder == null)
            {
                _loggingService.ShowError("Bitte wählen Sie zunächst das Hauptverzeichnis aus.");
                return false;
            }

            string? validatedKidName = _validator.IsKidNameValid();
            if (validatedKidName == null)
            {
                return false;
            }
            kidName = ConvertToTitleCase(validatedKidName);

            int? reportYearNullable = _validator.IsReportYearValid();
            if (!reportYearNullable.HasValue)
            {
                _loggingService.ShowError("Ungültiges Jahr.");
                return false;
            }
            reportYear = reportYearNullable.Value;

            return true;
        }


        private bool PrepareFilePaths(string kidName, int reportYear, out string sourceFolderPath, out string targetFolderPath, out string numericProtokollNumber)
        {
            sourceFolderPath = string.Empty;
            targetFolderPath = string.Empty;
            numericProtokollNumber = string.Empty;

            string group = ConvertToTitleCase(groupDropdown.Text);
            var (months, _) = ExtractMonthsFromExcel(group, kidName.Split(' ')[1], kidName.Split(' ')[0]);
            if (!months.HasValue)
            {
                _loggingService.ShowError("Fehler beim Extrahieren der Monate aus Excel.");
                return false;
            }

            var protokollbogenData = _validator.DetermineProtokollbogen(months.Value);
            if (!protokollbogenData.HasValue)
            {
                return false;
            }

            if (HomeFolder == null)
            {
                _loggingService.ShowError("Home folder is missing.");
                return false;
            }

            if (!protokollbogenData.HasValue)
            {
                _loggingService.ShowError("Protokollbogen data is missing.");
                return false;
            }

            sourceFolderPath = Path.Combine(HomeFolder.TrimEnd('\\'), protokollbogenData.Value.directoryPath.TrimStart('\\'));

            if (_fileManager == null)
            {
                _loggingService.ShowError("Der Dateimanager ist nicht initialisiert.");
                return false;
            }

            targetFolderPath = _fileManager.GetTargetPath(group, kidName, reportYear.ToString());

            if (protokollbogenData.HasValue)
            {
                numericProtokollNumber = ExtractProtokollNumber(protokollbogenData.Value.fileName + ".pdf");
                if (string.IsNullOrEmpty(numericProtokollNumber))
                {
                    _loggingService.ShowError("Fehler beim Extrahieren der Protokollnummer.");
                    return false;
                }
            }

            return true;
        }

        private void ExecuteFileOperations(string kidName, int reportYear, string sourceFolderPath, string targetFolderPath, string numericProtokollNumber)
        {
            bool isProtokollbogenChecked = protokollbogenAutoCheckbox.IsChecked == true;
            bool isAllgemeinerChecked = allgemeinerEntwicklungsberichtCheckbox.IsChecked == true;
            bool isVorschulChecked = vorschulentwicklungsberichtCheckbox.IsChecked == true;

            CopyFileIfChecked(isProtokollbogenChecked, Path.Combine(sourceFolderPath, numericProtokollNumber + ".pdf"), Path.Combine(targetFolderPath, numericProtokollNumber + ".pdf"));

            if (HomeFolder == null || string.IsNullOrEmpty("Entwicklungsboegen") || string.IsNullOrEmpty("Allgemeiner-Entwicklungsbericht.pdf"))
            {
                _loggingService.ShowError("One of the path components is missing.");
                return;
            }
            string allgemeinerFilePath = Path.Combine(HomeFolder, "Entwicklungsboegen", "Allgemeiner-Entwicklungsbericht.pdf");
            CopyFileIfChecked(isAllgemeinerChecked && File.Exists(allgemeinerFilePath), allgemeinerFilePath, Path.Combine(targetFolderPath, Path.GetFileName(allgemeinerFilePath)));

            string vorschulFilePath = Path.Combine(HomeFolder, "Entwicklungsboegen", "Vorschul-Entwicklungsbericht.pdf");
            CopyFileIfChecked(isVorschulChecked && File.Exists(vorschulFilePath), vorschulFilePath, Path.Combine(targetFolderPath, Path.GetFileName(vorschulFilePath)));

            if (_fileManager == null)
            {
                _loggingService.ShowError("Der Dateimanager ist nicht initialisiert.");
                return;
            }
            _fileManager.RenameFilesInTargetDirectory(targetFolderPath, kidName, ConvertToTitleCase(reportMonthDropdown.Text!), reportYear.ToString(), isAllgemeinerChecked, isVorschulChecked, isProtokollbogenChecked, numericProtokollNumber);

            MessageBox.Show("Dateien erfolgreich kopiert und umbenannt.", "Erfolgreich", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void CopyFileIfChecked(bool isChecked, string sourceFile, string targetFile)
        {
            if (isChecked)
            {
                if (_fileManager == null)
                {
                    _loggingService.ShowError("Der Dateimanager ist nicht initialisiert.");
                    return;
                }
                _fileManager.SafeCopyFile(sourceFile, targetFile);
            }
        }

        private bool AreAllRequiredFieldsFilled()
        {
            string selectedGroup = groupDropdown.Text;
            string childName = kidNameComboBox.Text;
            string selectedReportMonth = reportMonthDropdown.Text;
            string selectedReportYear = reportYearTextbox.Text;

            if (string.IsNullOrWhiteSpace(childName) || !childName.Contains(" "))
            {
                _loggingService.ShowError("Bitte geben Sie einen gültigen Namen mit Vor- und Nachnamen an.");
                return false;
            }

            if (string.IsNullOrWhiteSpace(selectedGroup) || string.IsNullOrWhiteSpace(selectedReportMonth) || string.IsNullOrWhiteSpace(selectedReportYear))
            {
                _loggingService.ShowError("Bitte füllen Sie alle geforderten Felder aus.");
                return false;
            }

            return true;
        }

        public void OnGroupSelected(object sender, SelectionChangedEventArgs e)
        {
            if (!_isWindowLoaded) return;
            if (string.IsNullOrEmpty(HomeFolder))
            {
                MessageBoxResult result = MessageBox.Show("Möchten Sie das Hauptverzeichnis ändern?", "Hauptverzeichnis nicht festgelegt", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    SelectHomeFolder();
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
                _allKidNames = kidNames;
                kidNameComboBox.ItemsSource = _allKidNames;
            }
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

        public void SelectHomeFolder()
        {
            Log.Information("SelectHomeFolder method called");
            var dialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            {
                dialog.Description = "Wählen Sie das Hauptverzeichnis aus";
                dialog.UseDescriptionForTitle = true;

                var result = dialog.ShowDialog();
                if (result.HasValue && result.Value)
                {
                    HomeFolder = dialog.SelectedPath;

                    if (!ValidationHelper.IsValidDirectoryPath(HomeFolder))
                    {
                        _loggingService.LogAndShowError($"Invalid home folder path: {HomeFolder}", "Der ausgewählte Pfad ist ungültig. Bitte wählen Sie ein gültiges Verzeichnis.");
                        return;
                    }
                    InitializeFileManager();

                    var settings = new AppSettings
                    {
                        HomeFolderPath = HomeFolder
                    };
                    settings.SaveSettings(settings);

                    MessageBox.Show($"Ausgewähltes Hauptverzeichnis: {HomeFolder}", "Hauptverzeichnis ausgewählt", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        private void OnSelectHomeFolderButtonClicked(object sender, RoutedEventArgs e)
        {
            SelectHomeFolder();
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            Log.CloseAndFlush();
        }
    }
}
