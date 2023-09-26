using ClosedXML.Excel;
using MahApps.Metro.Controls;
using Serilog;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static Automatisiertes_Kopieren.FileManager.StringUtilities;
using static Automatisiertes_Kopieren.LoggingService;

namespace Automatisiertes_Kopieren
{
    public partial class MainWindow : MetroWindow
    {
        private readonly static LoggingService _loggingService = new LoggingService();
        private FileManager? _fileManager;
        private int? _selectedProtokollbogenMonth;
        private bool _isHandlingCheckboxEvent = false;
        private int _previousGroupSelectionIndex = 0;
        private List<string> _allKidNames = new List<string>();

        public MainWindow()
        {
            _loggingService.InitializeLogger();
            InitializeComponent();
            var settings = new AppSettings().LoadSettings();
            if (settings != null && !string.IsNullOrEmpty(settings.HomeFolderPath))
            {
                homeFolder = settings.HomeFolderPath;
            }
            else
            {
                SelectHomeFolder();
            }

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
                    var defaultKidNames = GetKidNamesForGroup("Bären");
                    kidNameComboBox.ItemsSource = defaultKidNames;
                }
            }
        }

        private (double? months, string? error) ExtractMonthsFromExcel(string group, string lastName, string firstName)
        {
            if (string.IsNullOrEmpty(homeFolder))
            {
                _loggingService.ShowMessage("Bitte setzen Sie zuerst den Heimordner.", MessageType.Error);
                return (null, "HomeFolderNotSet");
            }
            string convertedGroupName = ConvertSpecialCharacters(group, ConversionType.Umlaute);
            string shortGroupName = convertedGroupName.Split(' ')[0];
            string filePath = $@"{homeFolder}\Entwicklungsberichte\{convertedGroupName} Entwicklungsberichte\Monatsrechner-Kinder-Zielsetzung-{shortGroupName}.xlsm";

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

                        if (double.TryParse(monthsValueRaw.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedValue))
                        {
                            return (Math.Round(parsedValue, 2), null);
                        }
                    }
                }
            }
            catch (FileNotFoundException)
            {
                _loggingService.LogAndShowMessage($"Die Datei {filePath} wurde nicht gefunden.",
                                                  "Die Datei wurde nicht gefunden. Bitte überprüfen Sie den Pfad.");
                return (null, "FileNotFound");
            }
            catch (IOException ioEx) when (ioEx.Message.Contains("because it is being used by another process"))
            {
                _loggingService.LogAndShowMessage($"Die Datei {filePath} wird von einem anderen Prozess verwendet.",
                                                  "Die Excel-Datei ist geöffnet. Bitte schließen Sie die Datei und versuchen Sie es erneut.");
                return (null, "FileInUse");
            }
            catch (Exception ex)
            {
                _loggingService.LogAndShowMessage($"Beim Verarbeiten der Excel-Datei ist ein unerwarteter Fehler aufgetreten: {ex.Message}",
                                                  "Ein unerwarteter Fehler ist aufgetreten. Bitte versuchen Sie es später erneut.");
                return (null, "UnexpectedError");
            }

            _loggingService.LogAndShowMessage($"Es konnte kein gültiger Monatswert für {firstName} {lastName} extrahiert werden.",
                                              "Es konnte kein gültiger Monatswert extrahiert werden. Bitte überprüfen Sie die Daten.");
            return (null, "ExtractionError");
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
                var nameParts = kidName.Split(' ');
                string kidFirstName = nameParts[0];
                string kidLastName = nameParts.Length > 1 ? nameParts[1] : "";

                var result = ExtractMonthsFromExcel(group, kidLastName, kidFirstName);
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

        private string ExtractProtokollNumber(string fileName)
        {
            var match = Regex.Match(fileName, @"Kind_Protokollbogen_(\d+)_Monate\.pdf");
            return match.Success ? match.Groups[1].Value + "_Monate" : string.Empty;
        }

        public static class OperationState
        {
            public static bool OperationsSuccessful { get; set; } = true;
        }


        private void PerformFileOperations()
        {
            OperationState.OperationsSuccessful = true;
            string sourceFolderPath = string.Empty;
            (string directoryPath, string fileName)? protokollbogenData = null;
            string numericProtokollNumber = string.Empty;

            string group = ConvertToTitleCase(groupDropdown.Text);
            if (homeFolder == null)
            {
                _loggingService.ShowMessage("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", MessageType.Error);
                return;
            }

            string? validatedKidName = ValidationHelper.ValidateKidName(kidNameComboBox.Text, homeFolder, groupDropdown.Text);
            if (validatedKidName == null)
            {
                _loggingService.ShowMessage("Ungültiger Kinder-Name.", MessageType.Error);
                return;
            }

            string kidName = ConvertToTitleCase(validatedKidName);
            string reportMonth = ConvertToTitleCase(reportMonthDropdown.Text);
            int? reportYearNullable = ValidationHelper.ValidateReportYearFromTextbox(reportYearTextbox.Text);
            if (!reportYearNullable.HasValue)
            {
                _loggingService.ShowMessage("Ungültiges Jahr.", MessageType.Error);
                return;
            }
            int reportYear = reportYearNullable.Value;

            var nameParts = kidName.Split(' ');
            string kidFirstName = nameParts[0];
            string kidLastName = nameParts[1];

            var (months, error) = ExtractMonthsFromExcel(group, kidLastName, kidFirstName);

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

            if (_fileManager == null)
            {
                _loggingService.ShowMessage("Der Dateimanager ist nicht initialisiert.", MessageType.Error);
                return;
            }
            string targetFolderPath = _fileManager.GetTargetPath(group, kidName, reportYear.ToString(), reportMonth);

            bool isAllgemeinerChecked = allgemeinerEntwicklungsberichtCheckbox.IsChecked == true;
            bool isVorschulChecked = vorschulentwicklungsberichtCheckbox.IsChecked == true;
            bool isProtokollbogenChecked = protokollbogenAutoCheckbox.IsChecked == true;

            if (protokollbogenData.HasValue && !string.IsNullOrEmpty(sourceFolderPath))
            {
                if (isProtokollbogenChecked)
                {
                    _fileManager.CopyFilesFromSourceToTarget(Path.Combine(sourceFolderPath, protokollbogenData.Value.fileName + ".pdf"), targetFolderPath, protokollbogenData.Value.fileName + ".pdf");
                }
            }

            string allgemeinerFilePath = Path.Combine(homeFolder, "Entwicklungsboegen", "Allgemeiner-Entwicklungsbericht.pdf");

            if (isAllgemeinerChecked && File.Exists(allgemeinerFilePath))
            {
                _fileManager.CopyFilesFromSourceToTarget(allgemeinerFilePath, targetFolderPath, Path.GetFileName(allgemeinerFilePath));
            }
            else if (!File.Exists(allgemeinerFilePath))
            {
                _loggingService.LogMessage($"File 'Allgemeiner-Entwicklungsbericht.pdf' not found at {allgemeinerFilePath}.", LogLevel.Warning);
            }

            string vorschulFilePath = Path.Combine(homeFolder, "Entwicklungsboegen", "Vorschul-Entwicklungsbericht.pdf");

            if (isVorschulChecked && File.Exists(vorschulFilePath))
            {
                _fileManager.CopyFilesFromSourceToTarget(vorschulFilePath, targetFolderPath, Path.GetFileName(vorschulFilePath));
            }
            else if (!File.Exists(vorschulFilePath))
            {
                _loggingService.LogMessage($"File 'Vorschul-Entwicklungsbericht.pdf' not found at {vorschulFilePath}.", LogLevel.Warning);
            }

            string protokollElterngespraechFilePath = Path.Combine(homeFolder, "Entwicklungsboegen", "Protokoll-Elterngespraech.pdf");

            if (File.Exists(protokollElterngespraechFilePath))
            {
                _fileManager.CopyFilesFromSourceToTarget(protokollElterngespraechFilePath, targetFolderPath, Path.GetFileName(protokollElterngespraechFilePath));
            }
            else
            {
                _loggingService.LogMessage($"File 'Protokoll-Elterngespraech.pdf' not found at {protokollElterngespraechFilePath}.", LogLevel.Warning);
            }


            if (protokollbogenData.HasValue)
            {
                numericProtokollNumber = ExtractProtokollNumber(protokollbogenData.Value.fileName + ".pdf");

                if (string.IsNullOrEmpty(numericProtokollNumber))
                {
                    _loggingService.ShowMessage("Fehler beim Extrahieren der Protokollnummer.", MessageType.Error);
                    return;
                }
            }

            _fileManager.RenameFilesInTargetDirectory(targetFolderPath, kidName, reportMonth, reportYear.ToString(), isAllgemeinerChecked, isVorschulChecked, isProtokollbogenChecked, numericProtokollNumber);
            
            if (OperationState.OperationsSuccessful)
            {
                _loggingService.ShowMessage("Dateien erfolgreich kopiert und umbenannt.", MessageType.Information);
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

        private List<string> GetKidNamesFromDirectory(string groupPath)
        {
            if (homeFolder != null)
            {
                string fullPath = Path.Combine(homeFolder, groupPath);
                if (Directory.Exists(fullPath))
                {
                    var directories = Directory.GetDirectories(fullPath);
                    return directories.Select(Path.GetFileName).OfType<string>().ToList();
                }
                else
                {
                    _loggingService.LogMessage($"Directory does not exist: {fullPath}", LogLevel.Warning);
                }
            }
            else
            {
                _loggingService.LogMessage("_homeFolder is not set.", LogLevel.Warning);
            }
            return new List<string>();
        }

        private List<string> GetKidNamesForGroup(string groupName)
        {
            string path = string.Empty;
            switch (groupName)
            {
                case "Bären":
                    path = "Entwicklungsberichte\\Baeren Entwicklungsberichte\\Aktuell";
                    break;
                case "Löwen":
                    path = "Entwicklungsberichte\\Loewen Entwicklungsberichte\\Aktuell";
                    break;
                case "Schnecken":
                    path = "Entwicklungsberichte\\Schnecken Beobachtungsberichte\\Aktuell";
                    break;
            }
            return GetKidNamesFromDirectory(path);
        }

        private void GroupDropdown_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? selectedGroup = groupDropdown.SelectedItem?.ToString();

            if (!string.IsNullOrEmpty(selectedGroup))
            {
                List<string> kidNames = GetKidNamesForGroup(selectedGroup);

                kidNameComboBox.ItemsSource = kidNames;
            }
        }

        private void OnKidNameComboBoxLoaded(object sender, RoutedEventArgs e)
        {
            var textBox = kidNameComboBox.Template.FindName("PART_EditableTextBox", kidNameComboBox) as TextBox;
            if (textBox != null)
            {
                textBox.TextChanged += OnKidNameComboBoxTextChanged;
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (groupDropdown.SelectedIndex == 0)
                {
                    OnGroupSelected(groupDropdown, new SelectionChangedEventArgs(ComboBox.SelectionChangedEvent, new List<object>(), new List<object> { groupDropdown.SelectedItem }));
                }
            }), System.Windows.Threading.DispatcherPriority.ContextIdle);
        }

        private void OnKidNameComboBoxPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (kidNameComboBox == null) return;

            var textBox = kidNameComboBox.Template.FindName("PART_EditableTextBox", kidNameComboBox) as TextBox;
            if (textBox == null) return;

            string futureText = textBox.Text.Insert(textBox.CaretIndex, e.Text);

            var filteredNames = _allKidNames.Where(name => name.StartsWith(futureText, StringComparison.OrdinalIgnoreCase)).ToList();

            if (filteredNames.Count == 0)
            {
                kidNameComboBox.ItemsSource = _allKidNames;
                kidNameComboBox.IsDropDownOpen = false;
                return;
            }

            kidNameComboBox.ItemsSource = filteredNames;
            kidNameComboBox.Text = futureText;
            textBox.CaretIndex = futureText.Length;
            kidNameComboBox.IsDropDownOpen = true;

            e.Handled = true;
        }

        private void OnKidNameComboBoxPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down || e.Key == Key.Up)
            {
                if (kidNameComboBox.IsDropDownOpen)
                {
                    e.Handled = false;
                }
            }
            else if (e.Key == Key.Enter)
            {
                if (kidNameComboBox.IsDropDownOpen)
                {
                    kidNameComboBox.SelectedItem = kidNameComboBox.Items.CurrentItem;
                    kidNameComboBox.IsDropDownOpen = false;
                }
            }
        }

        private bool _isUpdatingComboBox = false;

        private void OnKidNameComboBoxTextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingComboBox) return;
            if (kidNameComboBox == null) return;

            _isUpdatingComboBox = true;

            string input = kidNameComboBox.Text;

            var filteredNames = _allKidNames.Where(name => name.StartsWith(input, StringComparison.OrdinalIgnoreCase)).ToList();

            kidNameComboBox.ItemsSource = filteredNames.Count > 0 ? filteredNames : _allKidNames;
            kidNameComboBox.Text = input;
            kidNameComboBox.IsDropDownOpen = filteredNames.Count > 0;

            var textBox = kidNameComboBox.Template.FindName("PART_EditableTextBox", kidNameComboBox) as TextBox;
            if (textBox != null)
            {
                textBox.SelectionStart = input.Length;
            }

            _isUpdatingComboBox = false;
        }

        private void KidNameComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            if (groupDropdown.SelectedIndex == 0)
            {
                var defaultKidNames = GetKidNamesForGroup("Bären");
                _allKidNames = defaultKidNames;
                kidNameComboBox.ItemsSource = _allKidNames;
            }
        }

        private void OnGroupSelected(object sender, SelectionChangedEventArgs e)
        {
            if (kidNameComboBox == null)
            {
                _loggingService.LogMessage("kidNameComboBox is null.", LogLevel.Warning);
                return;
            }

            _loggingService.LogMessage($"_homeFolder value: {homeFolder}", LogLevel.Information);

            if (string.IsNullOrEmpty(homeFolder))
            {
                MessageBoxResult result = _loggingService.ShowMessage("Möchten Sie das Hauptverzeichnis ändern?", MessageType.Information, "Hauptverzeichnis nicht festgelegt", MessageBoxButton.YesNo);

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
                var kidNames = GetKidNamesForGroup(selectedGroup);
                _allKidNames = GetKidNamesForGroup(selectedGroup);
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

                    _loggingService.ShowMessage($"Ausgewähltes Hauptverzeichnis: {homeFolder}", MessageType.Information);
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

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            Log.CloseAndFlush();
        }
    }
}
