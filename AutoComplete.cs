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

        private static readonly Dictionary<string, string> GroupPaths = new Dictionary<string, string>
        {
            { "Bären", "Entwicklungsberichte\\Baeren Entwicklungsberichte\\Aktuell" },
            { "Löwen", "Entwicklungsberichte\\Loewen Entwicklungsberichte\\Aktuell" },
            { "Schnecken", "Entwicklungsberichte\\Schnecken Beobachtungsberichte\\Aktuell" }
        };

        public AutoComplete(MainWindow mainWindow)
        {
            _mainWindow = mainWindow;
        }

        public void OnKidNameComboBoxLoaded(object sender, RoutedEventArgs e)
        {
            var textBox = GetEditableTextBoxFromComboBox();
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

            var textBox = GetEditableTextBoxFromComboBox();
            if (textBox == null) return;

            string futureText = textBox.Text.Insert(textBox.CaretIndex, e.Text);

            FilterAndDisplayKidNames(futureText);

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

            FilterAndDisplayKidNames(_mainWindow.kidNameComboBox.Text);

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
            if (GroupPaths.TryGetValue(groupName, out var path))
            {
                Log.Information($"Constructed Path for {groupName}: {path}");
                return GetKidNamesFromDirectory(path);
            }
            else
            {
                Log.Warning($"Invalid group name: {groupName}");
                return new List<string>();
            }
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

        private TextBox? GetEditableTextBoxFromComboBox()
        {
            return _mainWindow.kidNameComboBox.Template.FindName("PART_EditableTextBox", _mainWindow.kidNameComboBox) as TextBox;
        }

        private void FilterAndDisplayKidNames(string input)
        {
            var filteredNames = _allKidNames.Where(name => name.StartsWith(input, StringComparison.OrdinalIgnoreCase)).ToList();

            _mainWindow.kidNameComboBox.ItemsSource = filteredNames.Count > 0 ? filteredNames : _allKidNames;
            _mainWindow.kidNameComboBox.Text = input;
            _mainWindow.kidNameComboBox.IsDropDownOpen = filteredNames.Count > 0;

            var textBox = GetEditableTextBoxFromComboBox();
            if (textBox != null)
            {
                textBox.SelectionStart = input.Length;
            }
        }
    }
}
