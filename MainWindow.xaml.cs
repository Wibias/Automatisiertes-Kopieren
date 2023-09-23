using ClosedXML.Excel;
using MahApps.Metro.Controls;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Configuration;
using System.Linq;
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
        private int? _selectedProtokollbogenMonth;
        private bool _isHandlingCheckboxEvent = false;
        private int _previousGroupSelectionIndex = 0;
        private List<string> _allKidNames = new List<string>();

        public MainWindow()
        {
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File("log-.txt", rollingInterval: RollingInterval.Day)
                .CreateLogger();
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
                return (null, "HomeFolderNotSet");
            }
            string convertedGroupName = ConvertSpecialCharacters(group, ConversionType.Umlaute);
            string shortGroupName = convertedGroupName.Split(' ')[0];
            string filePath = $@"{homeFolder}\Entwicklungsberichte\{convertedGroupName} Entwicklungsberichte\Monatsrechner-Kinder-Zielsetzung-{shortGroupName}.xlsm";

            if (!ValidationHelper.IsValidPath(filePath))
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
                Log.Error($"Die Datei {filePath} wurde nicht gefunden.");
                return (null, "FileNotFound");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Beim Verarbeiten der Excel-Datei ist ein unerwarteter Fehler aufgetreten.");
                return (null, "UnexpectedError");
            }

            Log.Error($"Es konnte kein gültiger Monatswert für {firstName} {lastName} extrahiert werden.");
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
                    MessageBox.Show("Bitte setzen Sie zuerst den Heimordner.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    protokollbogenAutoCheckbox.IsChecked = false;
                    return;
                }
                else if (result.error == "FileNotFound")
                {
                    MessageBox.Show("Das erforderliche Excel-Dokument konnte nicht gefunden werden. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    protokollbogenAutoCheckbox.IsChecked = false;
                    return;
                }
                else if (!result.months.HasValue)
                {
                    MessageBox.Show("Das Alter des Kindes konnte nicht aus Excel extrahiert werden.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
            // Input validation
            if (!IsValidInput())
            {
                // If the input is not valid, return and avoid performing the operations
                return;
            }

            // Perform the required file operations if input is valid
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
                    MessageBox.Show("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }

            // Check if _homeFolder is null
            if (homeFolder == null)
            {
                MessageBox.Show("Das Hauptverzeichnis ist nicht festgelegt.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            // Validate kid's name
            string kidName = kidNameComboBox.Text;
            string? validatedKidName = ValidationHelper.ValidateKidName(kidName, homeFolder, groupDropdown.Text);
            if (string.IsNullOrEmpty(validatedKidName))
            {
                // Stop processing because the name wasn't valid or another error occurred.
                MessageBox.Show("Ungültiger Kinder-Name", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            // Validate report year
            string reportYearText = reportYearTextbox.Text;
            try
            {
                int? parsedYear = ValidationHelper.ValidateReportYearFromTextbox(reportYearText);
                if (!parsedYear.HasValue)
                {
                    MessageBox.Show("Bitte geben Sie ein gültiges Jahr für den Bericht an.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        private string ExtractProtokollNumber(string fileName)
        {
            var match = Regex.Match(fileName, @"Kind_Protokollbogen_(\d+)_Monate\.pdf");
            return match.Success ? match.Groups[1].Value + "_Monate" : string.Empty;
        }

        private void PerformFileOperations()
        {
            string sourceFolderPath = string.Empty;
            (string directoryPath, string fileName)? protokollbogenData = null;
            string numericProtokollNumber = string.Empty;

            string group = FileManager.StringUtilities.ConvertToTitleCase(groupDropdown.Text);
            if (homeFolder == null)
            {
                MessageBox.Show("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string? validatedKidName = ValidationHelper.ValidateKidName(kidNameComboBox.Text, homeFolder, groupDropdown.Text);
            if (validatedKidName == null)
            {
                MessageBox.Show("Ungültiger Kinder-Name.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string kidName = FileManager.StringUtilities.ConvertToTitleCase(validatedKidName);
            string reportMonth = FileManager.StringUtilities.ConvertToTitleCase(reportMonthDropdown.Text);
            int? reportYearNullable = ValidationHelper.ValidateReportYearFromTextbox(reportYearTextbox.Text);
            if (!reportYearNullable.HasValue)
            {
                MessageBox.Show("Ungültiges Jahr.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            int reportYear = reportYearNullable.Value;

            var nameParts = kidName.Split(' ');
            string kidFirstName = nameParts[0];
            string kidLastName = nameParts[1];

            var (months, error) = ExtractMonthsFromExcel(group, kidLastName, kidFirstName);

            if (months.HasValue)
            {
                var protokollbogenResult = ValidationHelper.DetermineProtokollbogen(months.Value);
                if (protokollbogenResult.HasValue)
                {
                    protokollbogenData = protokollbogenResult;

                    if (homeFolder == null)
                    {
                        MessageBox.Show("Das Hauptverzeichnis ist nicht festgelegt.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    string cleanedHomeFolder = homeFolder.TrimEnd('\\');
                    string cleanedDirectoryPath = protokollbogenData.Value.directoryPath.TrimStart('\\');

                    sourceFolderPath = Path.Combine(cleanedHomeFolder, cleanedDirectoryPath);
                }
            }

            else
            {
                MessageBox.Show("Fehler beim Extrahieren der Monate aus Excel.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (_fileManager == null)
            {
                MessageBox.Show("Der Dateimanager ist nicht initialisiert.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            string targetFolderPath = _fileManager.GetTargetPath(group, kidName, reportYear.ToString());

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

            if (!ValidationHelper.IsValidPath(allgemeinerFilePath))
            {
                Log.Error($"Verzeichnis existiert nicht: {allgemeinerFilePath}");
                return;
            }

            if (isAllgemeinerChecked && File.Exists(allgemeinerFilePath))
            {
                _fileManager.CopyFilesFromSourceToTarget(allgemeinerFilePath, targetFolderPath, Path.GetFileName(allgemeinerFilePath));
            }
            else if (!File.Exists(allgemeinerFilePath))
            {
                Log.Warning($"File 'Allgemeiner-Entwicklungsbericht.pdf' not found at {allgemeinerFilePath}.");
            }

            string vorschulFilePath = Path.Combine(homeFolder, "Entwicklungsboegen", "Vorschul-Entwicklungsbericht.pdf");

            if (!ValidationHelper.IsValidPath(vorschulFilePath))
            {
                Log.Error($"Verzeichnis existiert nicht: {vorschulFilePath}");
                return;
            }

            if (isVorschulChecked && File.Exists(vorschulFilePath))
            {
                _fileManager.CopyFilesFromSourceToTarget(vorschulFilePath, targetFolderPath, Path.GetFileName(vorschulFilePath));
            }
            else if (!File.Exists(vorschulFilePath))
            {
                Log.Warning($"Datei 'Vorschul-Entwicklungsbericht.pdf' nicht gefunden unter {vorschulFilePath}.");
            }

            if (protokollbogenData.HasValue)
            {
                numericProtokollNumber = ExtractProtokollNumber(protokollbogenData.Value.fileName + ".pdf");

                if (string.IsNullOrEmpty(numericProtokollNumber))
                {
                    MessageBox.Show("Fehler beim Extrahieren der Protokollnummer.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            _fileManager.RenameFilesInTargetDirectory(targetFolderPath, kidName, reportMonth, reportYear.ToString(), isAllgemeinerChecked, isVorschulChecked, isProtokollbogenChecked, numericProtokollNumber);

            MessageBox.Show("Dateien erfolgreich kopiert und umbenannt.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
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
                MessageBox.Show("Bitte geben Sie einen gültigen Namen mit Vor- und Nachnamen an.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(selectedGroup) || string.IsNullOrWhiteSpace(selectedReportMonth) || string.IsNullOrWhiteSpace(selectedReportYear))
            {
                MessageBox.Show("Bitte füllen Sie alle geforderten Felder aus.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        private List<string> GetKidNamesFromDirectory(string groupPath)
        {
            if (homeFolder != null)
            {
                string fullPath = Path.Combine(homeFolder, groupPath);

                if (!ValidationHelper.IsValidPath(fullPath))
                {
                    Log.Error($"Verzeichnis existiert nicht: {fullPath}");
                    return new List<string>();
                };

                if (Directory.Exists(fullPath))
                {
                    var directories = Directory.GetDirectories(fullPath);
                    return directories.Select(Path.GetFileName).OfType<string>().ToList();
                }
                else
                {
                    Log.Warning($"Verzeichnis existiert nicht: {fullPath}");
                }
            }
            else
            {
                Log.Warning("_homeFolder ist nicht gesetzt.");
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
            Log.Information($"Constructed Path for {groupName}: {path}");
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
            textBox.CaretIndex = futureText.Length;  // Set the caret position to the end of the input
            kidNameComboBox.IsDropDownOpen = true;

            // This is important to prevent the ComboBox from handling the input and selecting an item
            e.Handled = true;
        }

        private void OnKidNameComboBoxPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down || e.Key == Key.Up)
            {
                // If the dropdown is open and the user presses the arrow keys, let the ComboBox handle the navigation
                if (kidNameComboBox.IsDropDownOpen)
                {
                    e.Handled = false;
                }
            }
            else if (e.Key == Key.Enter)
            {
                // If the user presses Enter, select the highlighted item
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
                textBox.SelectionStart = input.Length;  // Set the caret position to the end of the input
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
            if (string.IsNullOrEmpty(homeFolder))
            {
                MessageBoxResult result = MessageBox.Show("Möchten Sie das Hauptverzeichnis ändern?", "Hauptverzeichnis nicht festgelegt", MessageBoxButton.YesNo, MessageBoxImage.Question);

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
                Log.Information($"e.AddedItems[0] type: {e.AddedItems[0]?.GetType().Name ?? "null"}, value: {e.AddedItems[0]}");
                Log.Information($"Selected group: {selectedGroup}");
                var kidNames = GetKidNamesForGroup(selectedGroup);
                _allKidNames = GetKidNamesForGroup(selectedGroup);
                kidNameComboBox.ItemsSource = _allKidNames;
            }
            else if (e.AddedItems.Count > 0)
            {
                Log.Warning($"e.AddedItems[0] type: {e.AddedItems[0]?.GetType().Name ?? "null"}, value: {e.AddedItems[0]}");
                Log.Warning("Selected group is empty or not a valid ComboBoxItem.");
            }
            else
            {
                Log.Warning("No group selected.");
            }
        }

        private void SelectHomeFolder()
        {
            Log.Information("SelectHomeFolder method called");
            var dialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            {
                dialog.Description = "Wählen Sie das Hauptverzeichnis aus";
                dialog.UseDescriptionForTitle = true;

                var result = dialog.ShowDialog();
                if (result.HasValue && result.Value)
                {
                    homeFolder = dialog.SelectedPath;

                    if (!ValidationHelper.IsValidPath(homeFolder))
                    {
                        Log.Error($"Invalid home folder path: {homeFolder}");
                        MessageBox.Show("The selected path is invalid. Please choose a valid directory.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                    InitializeFileManager();

                    var settings = new AppSettings
                    {
                        HomeFolderPath = homeFolder
                    };
                    settings.SaveSettings(settings);

                    MessageBox.Show($"Ausgewähltes Hauptverzeichnis: {homeFolder}", "Hauptverzeichnis ausgewählt", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        private void OnSelectHomeFolderButtonClicked(object sender, RoutedEventArgs e)
        {
            SelectHomeFolder();
        }

        private void InitializeFileManager()
        {
            if (homeFolder != null && ValidationHelper.IsValidPath(homeFolder))
            {
                _fileManager = new FileManager(homeFolder);
            }
            else
            {
                MessageBox.Show("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            Log.CloseAndFlush();
        }
    }
}
