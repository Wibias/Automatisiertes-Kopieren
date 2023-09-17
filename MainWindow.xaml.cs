using System.Windows;
using System.Windows.Forms;
using System.Linq;
using System;
using System.Collections.Generic;
using Serilog;
using ClosedXML.Excel;
using MessageBox = System.Windows.MessageBox;
using MahApps.Metro.Controls;
using System.Windows.Controls;

namespace Automatisiertes_Kopieren
{
    public partial class MainWindow : MetroWindow
    {
        private const int StartRow = 7;
        private const int EndRow = 31;
        private const string WorksheetName = "Monatsrechner";

        private string? _homeFolder;
        private FileManager? _fileManager;
        private int? _selectedProtokollbogenMonth;

        public MainWindow()
        {
            // Initialize Serilog
            Serilog.Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .CreateLogger();
            InitializeComponent();
            protokollbogenAutoCheckbox.Checked += OnProtokollbogenAutoCheckboxChanged;
            protokollbogenAutoCheckbox.Unchecked += OnProtokollbogenAutoCheckboxChanged;
        }

        private double? ExtractMonthsFromExcel(string group, string lastName, string firstName)
        {
            string shortGroupName = group.Split(' ')[0];
            string filePath = $@"{_homeFolder}\Entwicklungsberichte\Entwicklungsberichte\{group}\Monatsrechner-Kinder-Zielsetzung-{shortGroupName}.xlsm";

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
                        return parsedValue;
                    }
                }
            }

            Serilog.Log.Error($"Es konnte kein gültiger Monatswert für {firstName} {lastName} extrahiert werden.");
            return null;
        }

        private void OnProtokollbogenAutoCheckboxChanged(object sender, RoutedEventArgs e)
        {
            if (protokollbogenAutoCheckbox.IsChecked == true)
            {
                string group = groupDropdown.Text; // Assuming you have a dropdown for the group
                string kidName = kidNameTextbox.Text; // Assuming you have a textbox for the kid's name
                var nameParts = kidName.Split(' ');
                string kidFirstName = nameParts[0];
                string kidLastName = nameParts.Length > 1 ? nameParts[1] : "";

                double? childAgeInMonths = ExtractMonthsFromExcel(group, kidLastName, kidFirstName);
                if (!childAgeInMonths.HasValue)
                {
                    MessageBox.Show("Failed to extract the child's age from Excel.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                _selectedProtokollbogenMonth = (int)Math.Round(childAgeInMonths.Value); // Rounding to get the nearest whole month
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
            int? parsedYear = null;
            try
            {
                parsedYear = ValidationHelper.ValidateReportYearFromTextbox(reportYearText);
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
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Wählen Sie das Hauptverzeichnis aus";
                dialog.ShowNewFolderButton = true;

                DialogResult result = dialog.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
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
            Serilog.Log.CloseAndFlush();
        }

    }
}
