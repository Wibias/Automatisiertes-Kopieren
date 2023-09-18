using ClosedXML.Excel;
using MahApps.Metro.Controls;
using Serilog;
using System;
using System.IO;
using System.Windows;
using System.Windows.Forms;
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

        public MainWindow()
        {
            // Initialize Serilog
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File("log-.txt", rollingInterval: RollingInterval.Day)
                .CreateLogger();
            InitializeComponent();
            protokollbogenAutoCheckbox.Checked += OnProtokollbogenAutoCheckboxChanged;
            protokollbogenAutoCheckbox.Unchecked += OnProtokollbogenAutoCheckboxChanged;
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
                            continue;  // Skip this iteration if names don't match
                        }

                        var monthsValueRaw = worksheet.Cell(row, 6).Value.ToString();

                        if (double.TryParse(monthsValueRaw.Replace(",", "."), out double parsedValue))
                        {
                            return (parsedValue, null); // Return the parsed value and null for the error
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
                string kidName = kidNameTextbox.Text;
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
            string kidName = kidNameTextbox.Text;
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

        private void PerformFileOperations()
        {
            // Extract input values
            string group = FileManager.StringUtilities.ConvertToTitleCase(groupDropdown.Text);
            // Check if _homeFolder is null
            if (_homeFolder == null)
            {
                MessageBox.Show("Bitte wählen Sie zunächst das Hauptverzeichnis aus.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Validate kid's name and check the result
            string? validatedKidName = ValidationHelper.ValidateKidName(kidNameTextbox.Text, _homeFolder, groupDropdown.Text);
            if (validatedKidName == null)
            {
                MessageBox.Show("Ungültiger Kinder-Name.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Convert the validated kid's name to title case
            string kidName = FileManager.StringUtilities.ConvertToTitleCase(validatedKidName);
            string reportMonth = FileManager.StringUtilities.ConvertToTitleCase(reportMonthDropdown.Text);
            int? reportYearNullable = ValidationHelper.ValidateReportYearFromTextbox(reportYearTextbox.Text);
            if (!reportYearNullable.HasValue)
            {
                // Handle the error here: for example, show a message to the user
                MessageBox.Show("Ungültiges Jahr.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            int reportYear = reportYearNullable.Value;

            // Extract first and last names from the kid's name
            var nameParts = kidName.Split(' ');
            string kidFirstName = nameParts[0];
            string kidLastName = nameParts[1];

            // Define the source path for the files that need to be copied
            string sourceFolderPath = _homeFolder;
            if (_fileManager == null)
            {
                MessageBox.Show("Der Dateimanager ist nicht initialisiert.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            string targetFolderPath = _fileManager.GetTargetPath(group, kidName, reportMonth, reportYear.ToString());


            // Copy and rename files
            _fileManager.CopyFilesFromSourceToTarget(sourceFolderPath, targetFolderPath);

            bool isAllgemeinerChecked = allgemeinerEntwicklungsberichtCheckbox.IsChecked == true;
            bool isVorschulentwicklungsberichtChecked = vorschulentwicklungsberichtCheckbox.IsChecked == true;
            bool isProtokollbogenChecked = protokollbogenAutoCheckbox.IsChecked == true;
            if (!_selectedProtokollbogenMonth.HasValue)
            {
                MessageBox.Show("Protokollbogen Monat ist nicht gesetzt.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            int protokollNumber = _selectedProtokollbogenMonth.Value;

            _fileManager.RenameFilesInTargetDirectory(targetFolderPath, kidName, reportMonth, reportYear.ToString(), isAllgemeinerChecked, isVorschulentwicklungsberichtChecked, isProtokollbogenChecked, protokollNumber);

            // Provide feedback to the user
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
            string childName = kidNameTextbox.Text;
            string selectedReportMonth = reportMonthDropdown.Text;
            string selectedReportYear = reportYearTextbox.Text;

            if (string.IsNullOrWhiteSpace(childName) || !childName.Contains(" "))
            {
                MessageBox.Show("Bitte geben Sie einen gültigen Namen mit Vor- und Nachnamen an.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            // Consolidated the field checks
            if (string.IsNullOrWhiteSpace(selectedGroup) || string.IsNullOrWhiteSpace(selectedReportMonth) || string.IsNullOrWhiteSpace(selectedReportYear))
            {
                MessageBox.Show("Bitte füllen Sie alle geforderten Felder aus.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        private void SelectHomeFolder()
        {
            var dialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            {
                dialog.Description = "Wählen Sie das Hauptverzeichnis aus";
                dialog.UseDescriptionForTitle = true; // This will make the description appear as title

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
