using ClosedXML.Excel;
using MahApps.Metro.Controls;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using static Automatisiertes_Kopieren.FileManager.StringUtilities;
using MessageBox = System.Windows.MessageBox;

namespace Automatisiertes_Kopieren
{
    public partial class MainWindow : MetroWindow
    {

        private string? _homeFolder;
        private FileManager? _fileManager;
        private int? _selectedProtokollbogenMonth;
        private bool _isHandlingCheckboxEvent = false;
        private int _previousGroupSelectionIndex = 0;

        public MainWindow()
        {
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File("log-.txt", rollingInterval: RollingInterval.Day)
                .CreateLogger();
            InitializeComponent();

            protokollbogenAutoCheckbox.Checked += OnProtokollbogenAutoCheckboxChanged;
            protokollbogenAutoCheckbox.Unchecked += OnProtokollbogenAutoCheckboxChanged;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // For testing purposes, set the _homeFolder directly here
            _homeFolder = "G:\\Entwicklungsberichte"; // Replace with your actual folder path

            if (groupDropdown.SelectedIndex == 0 && !string.IsNullOrEmpty(_homeFolder))
            {
                var defaultKidNames = GetKidNamesForGroup("Bären");
                kidNameComboBox.ItemsSource = defaultKidNames;
            }
        }

        private (double? months, string? error) ExtractMonthsFromExcel(string group, string lastName, string firstName)
        {
            if (string.IsNullOrEmpty(_homeFolder))
            {
                return (null, "HomeFolderNotSet");
            }
            string convertedGroupName = ConvertSpecialCharacters(group, ConversionType.Umlaute);
            string shortGroupName = convertedGroupName.Split(' ')[0];
            string filePath = $@"{_homeFolder}\Entwicklungsberichte\{convertedGroupName} Entwicklungsberichte\Monatsrechner-Kinder-Zielsetzung-{shortGroupName}.xlsm";

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
            if (_isHandlingCheckboxEvent) return; // Exit if already handling the event

            _isHandlingCheckboxEvent = true; // Set the flag to true

            if (e.RoutedEvent.Name == "Checked")
            {
                // Handle the Checked logic here
                HandleProtokollbogenAutoCheckbox(true);
            }
            else if (e.RoutedEvent.Name == "Unchecked")
            {
                // Handle the Unchecked logic here
                HandleProtokollbogenAutoCheckbox(false);
            }

            _isHandlingCheckboxEvent = false; // Reset the flag
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
                    protokollbogenAutoCheckbox.IsChecked = false; // Uncheck the checkbox to prevent further processing
                    return;
                }
                else if (result.error == "FileNotFound")
                {
                    MessageBox.Show("Das erforderliche Excel-Dokument konnte nicht gefunden werden. Bitte überprüfen Sie den Pfad und versuchen Sie es erneut.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    protokollbogenAutoCheckbox.IsChecked = false; // Uncheck the checkbox to prevent further processing
                    return;
                }
                else if (!result.months.HasValue)
                {
                    MessageBox.Show("Das Alter des Kindes konnte nicht aus Excel extrahiert werden.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    protokollbogenAutoCheckbox.IsChecked = false; // Uncheck the checkbox to prevent further processing
                    return;
                }
                _selectedProtokollbogenMonth = (int)Math.Round(result.months.Value); // Rounding to get the nearest whole month
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
                if (_homeFolder != null)
                {
                    _fileManager = new FileManager(_homeFolder);
                }
                else
                {
                    MessageBox.Show("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }

            // Check if _homeFolder is null
            if (_homeFolder == null)
            {
                MessageBox.Show("Das Hauptverzeichnis ist nicht festgelegt.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            // Validate kid's name
            string kidName = kidNameComboBox.Text;
            string? validatedKidName = ValidationHelper.ValidateKidName(kidName, _homeFolder, groupDropdown.Text);
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
            if (_homeFolder == null)
            {
                MessageBox.Show("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string? validatedKidName = ValidationHelper.ValidateKidName(kidNameComboBox.Text, _homeFolder, groupDropdown.Text);
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

            // Extract months from Excel
            var (months, error) = ExtractMonthsFromExcel(group, kidLastName, kidFirstName);

            if (months.HasValue)
            {
                var protokollbogenResult = ValidationHelper.DetermineProtokollbogen(months.Value);
                if (protokollbogenResult.HasValue)
                {
                    protokollbogenData = protokollbogenResult;

                    if (_homeFolder == null)
                    {
                        MessageBox.Show("Das Hauptverzeichnis ist nicht festgelegt.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    string cleanedHomeFolder = _homeFolder.TrimEnd('\\');
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

            string allgemeinerFilePath = Path.Combine(_homeFolder, "Entwicklungsboegen", "Allgemeiner-Entwicklungsbericht.pdf");

            if (isAllgemeinerChecked && File.Exists(allgemeinerFilePath))
            {
                _fileManager.CopyFilesFromSourceToTarget(allgemeinerFilePath, targetFolderPath, Path.GetFileName(allgemeinerFilePath));
            }
            else if (!File.Exists(allgemeinerFilePath))
            {
                Log.Warning($"File 'Allgemeiner-Entwicklungsbericht.pdf' not found at {allgemeinerFilePath}.");
            }

            string vorschulFilePath = Path.Combine(_homeFolder, "Entwicklungsboegen", "Vorschul-Entwicklungsbericht.pdf");

            if (isVorschulChecked && File.Exists(vorschulFilePath))
            {
                _fileManager.CopyFilesFromSourceToTarget(vorschulFilePath, targetFolderPath, Path.GetFileName(vorschulFilePath));
            }
            else if (!File.Exists(vorschulFilePath))
            {
                Log.Warning($"File 'Vorschul-Entwicklungsbericht.pdf' not found at {vorschulFilePath}.");
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
            if (!string.IsNullOrEmpty(_homeFolder))
                return true;

            SelectHomeFolder();
            return !string.IsNullOrEmpty(_homeFolder);
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
            if (_homeFolder != null)
            {
                string fullPath = Path.Combine(_homeFolder, groupPath);
                Log.Information($"Full path: {fullPath}");

                if (Directory.Exists(fullPath))
                {
                    var directories = Directory.GetDirectories(fullPath);
                    Log.Information($"Found directories: {string.Join(", ", directories)}");
                    return directories.Select(Path.GetFileName).OfType<string>().ToList();
                }
                else
                {
                    Log.Warning($"Directory does not exist: {fullPath}");
                }
            }
            else
            {
                Log.Warning("_homeFolder is not set.");
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
            if (groupDropdown.SelectedIndex == 0)
            {
                var defaultKidNames = GetKidNamesForGroup("Bären");
                kidNameComboBox.ItemsSource = defaultKidNames;
            }
        }

        private void OnKidNameComboBoxTextChanged(object sender, TextChangedEventArgs e)
        {
            if (kidNameComboBox == null) return;

            string input = kidNameComboBox.Text;

            var allKidNames = kidNameComboBox.ItemsSource as List<string>;

            if (allKidNames == null) return;

            var filteredNames = allKidNames.Where(name => name.StartsWith(input, StringComparison.OrdinalIgnoreCase)).ToList();

            kidNameComboBox.ItemsSource = filteredNames;

            kidNameComboBox.Text = input;

            kidNameComboBox.IsDropDownOpen = true;
        }

        private void OnGroupSelected(object sender, SelectionChangedEventArgs e)
        {
            Log.Information("OnGroupSelected triggered");
            if (kidNameComboBox == null)
            {
                Log.Warning("kidNameComboBox is null.");
                return;
            }

            Log.Information($"_homeFolder value: {_homeFolder}");

            if (string.IsNullOrEmpty(_homeFolder))
            {
                MessageBox.Show("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                groupDropdown.SelectionChanged -= OnGroupSelected;

                groupDropdown.SelectedIndex = _previousGroupSelectionIndex;

                groupDropdown.SelectionChanged += OnGroupSelected;

                return;
            }

            _previousGroupSelectionIndex = groupDropdown.SelectedIndex;

            if (e.AddedItems.Count > 0)
            {
                Log.Information($"e.AddedItems[0] type: {e.AddedItems[0].GetType().Name}, value: {e.AddedItems[0]}");
                if (e.AddedItems.Count > 0 && e.AddedItems[0] is ComboBoxItem comboBoxItem && comboBoxItem.Content is string selectedGroup && !string.IsNullOrEmpty(selectedGroup))
                {
                    Log.Information($"Selected group: {selectedGroup}");
                    var kidNames = GetKidNamesForGroup(selectedGroup);
                    Log.Information($"Kid names for {selectedGroup}: {string.Join(", ", kidNames)}");
                    kidNameComboBox.ItemsSource = kidNames;
                }
                else
                {
                    Log.Warning("Selected group is empty or not a valid ComboBoxItem.");
                }
            }
            else
            {
                Log.Warning("No group selected.");
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
                    _homeFolder = dialog.SelectedPath;
                    InitializeFileManager();
                    MessageBox.Show($"Ausgewähltes Hauptverzeichnis: {_homeFolder}", "Hauptverzeichnis ausgewählt", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }
        private void OnSelectHomeFolderButtonClicked(object sender, RoutedEventArgs e)
        {
            SelectHomeFolder();
        }

        private void InitializeFileManager()
        {
            if (_homeFolder != null)
            {
                _fileManager = new FileManager(_homeFolder);
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
