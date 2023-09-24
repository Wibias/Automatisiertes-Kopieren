using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Automatisiertes_Kopieren
{
    public class AutoComplete
    {
        private List<string> _allKidNames = new List<string>();
        private MainWindow _mainWindow;

        public AutoComplete(MainWindow mainWindow)
        {
            _mainWindow = mainWindow;
        }

        public void OnKidNameComboBoxLoaded(object sender, RoutedEventArgs e)
        {
            var textBox = _mainWindow.kidNameComboBox.Template.FindName("PART_EditableTextBox", _mainWindow.kidNameComboBox) as TextBox;
            if (textBox != null)
            {
                textBox.TextChanged += OnKidNameComboBoxTextChanged;
            }

            _mainWindow.Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_mainWindow.groupDropdown.SelectedIndex == 0)
                {
                    _mainWindow.OnGroupSelected(_mainWindow.groupDropdown, new SelectionChangedEventArgs(ComboBox.SelectionChangedEvent, new List<object>(), new List<object> { _mainWindow.groupDropdown.SelectedItem }));
                }
            }), System.Windows.Threading.DispatcherPriority.ContextIdle);
        }

        public void OnKidNameComboBoxPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (_mainWindow.kidNameComboBox == null) return;

            var textBox = _mainWindow.kidNameComboBox.Template.FindName("PART_EditableTextBox", _mainWindow.kidNameComboBox) as TextBox;
            if (textBox == null) return;

            string futureText = textBox.Text.Insert(textBox.CaretIndex, e.Text);

            var filteredNames = _allKidNames.Where(name => name.StartsWith(futureText, StringComparison.OrdinalIgnoreCase)).ToList();

            if (filteredNames.Count == 0)
            {
                _mainWindow.kidNameComboBox.ItemsSource = _allKidNames;
                _mainWindow.kidNameComboBox.IsDropDownOpen = false;
                return;
            }

            _mainWindow.kidNameComboBox.ItemsSource = filteredNames;
            _mainWindow.kidNameComboBox.Text = futureText;
            textBox.CaretIndex = futureText.Length;
            _mainWindow.kidNameComboBox.IsDropDownOpen = true;

            e.Handled = true;
        }

        public void OnKidNameComboBoxPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down || e.Key == Key.Up)
            {
                if (_mainWindow.kidNameComboBox.IsDropDownOpen)
                {
                    e.Handled = false;
                }
            }
            else if (e.Key == Key.Enter)
            {
                if (_mainWindow.kidNameComboBox.IsDropDownOpen)
                {
                    _mainWindow.kidNameComboBox.SelectedItem = _mainWindow.kidNameComboBox.Items.CurrentItem;
                    _mainWindow.kidNameComboBox.IsDropDownOpen = false;
                }
            }
        }

        private bool _isUpdatingComboBox = false;

        public void OnKidNameComboBoxTextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingComboBox) return;
            if (_mainWindow.kidNameComboBox == null) return;

            _isUpdatingComboBox = true;

            string input = _mainWindow.kidNameComboBox.Text;

            var filteredNames = _allKidNames.Where(name => name.StartsWith(input, StringComparison.OrdinalIgnoreCase)).ToList();

            _mainWindow.kidNameComboBox.ItemsSource = filteredNames.Count > 0 ? filteredNames : _allKidNames;
            _mainWindow.kidNameComboBox.Text = input;
            _mainWindow.kidNameComboBox.IsDropDownOpen = filteredNames.Count > 0;

            var textBox = _mainWindow.kidNameComboBox.Template.FindName("PART_EditableTextBox", _mainWindow.kidNameComboBox) as TextBox;
            if (textBox != null)
            {
                textBox.SelectionStart = input.Length;
            }

            _isUpdatingComboBox = false;
        }

        public void KidNameComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            if (_mainWindow.groupDropdown.SelectedIndex == 0)
            {
                var defaultKidNames = GetKidNamesForGroup("Bären");
                _allKidNames = defaultKidNames;
                _mainWindow.kidNameComboBox.ItemsSource = _allKidNames;
            }
        }

        public List<string> GetKidNamesForGroup(string groupName)
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

        private List<string> GetKidNamesFromDirectory(string groupPath)
        {
            if (_mainWindow.HomeFolder != null)
            {
                string fullPath = Path.Combine(_mainWindow.HomeFolder, groupPath);

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
    }
}