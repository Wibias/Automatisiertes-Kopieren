using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using static Automatisiertes_Kopieren.LoggingHelper;
using ComboBox = System.Windows.Controls.ComboBox;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using TextBox = System.Windows.Controls.TextBox;

namespace Automatisiertes_Kopieren;

public class AutoCompleteHelper
{
    private readonly MainWindow _mainWindow;
    private List<string> _allKidNames = new();
    private bool _isArrowKeySelection;

    public AutoCompleteHelper(MainWindow mainWindow)
    {
        _mainWindow = mainWindow;
    }

    public void OnKidNameComboBoxPreviewTextInput(TextCompositionEventArgs e)
    {
        if (_mainWindow.KidNameComboBox?.Template.FindName("PART_EditableTextBox", _mainWindow.KidNameComboBox) is not
            TextBox textBox) return;

        var currentText = textBox.Text + e.Text;

        var filteredNames = _allKidNames
            .Where(name => name.StartsWith(currentText, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (filteredNames.Count == 0)
        {
            _mainWindow.KidNameComboBox.ItemsSource = _allKidNames;
            _mainWindow.KidNameComboBox.IsDropDownOpen = false;
            return;
        }

        _mainWindow.KidNameComboBox.ItemsSource = filteredNames;
        _mainWindow.KidNameComboBox.Text = currentText;
        textBox.CaretIndex = currentText.Length;
        _mainWindow.KidNameComboBox.IsDropDownOpen = true;
        _mainWindow.GroupDropdown.SelectedItem = "Bären";

        e.Handled = true;

        if (!_mainWindow.KidNameComboBox.IsDropDownOpen && _mainWindow.KidNameComboBox.HasItems)
            _mainWindow.KidNameComboBox.IsDropDownOpen = true;

        if (!_isArrowKeySelection) return;
        e.Handled = true;
        _isArrowKeySelection = false;
    }

    public void OnKidNameComboBoxPreviewKeyDown(KeyEventArgs e)
    {
        _isArrowKeySelection = false;

        switch (e.Key)
        {
            case Key.Down:
                if (!_mainWindow.KidNameComboBox.IsDropDownOpen && _mainWindow.KidNameComboBox.HasItems)
                {
                    _mainWindow.KidNameComboBox.IsDropDownOpen = true;
                    _mainWindow.KidNameComboBox.SelectedIndex = -1;
                }
                else if (_mainWindow.KidNameComboBox.IsDropDownOpen)
                {
                    _isArrowKeySelection = true;

                    if (_mainWindow.KidNameComboBox.SelectedIndex == -1)
                    {
                        _mainWindow.KidNameComboBox.SelectedIndex = 0;
                    }
                    else
                    {
                        var nextIndex = _mainWindow.KidNameComboBox.SelectedIndex + 1;
                        if (nextIndex < _mainWindow.KidNameComboBox.Items.Count)
                            _mainWindow.KidNameComboBox.SelectedIndex = nextIndex;
                    }

                    e.Handled = true;
                }

                break;

            case Key.Up:
                if (_mainWindow.KidNameComboBox.IsDropDownOpen && _mainWindow.KidNameComboBox.SelectedIndex > 0)
                {
                    _isArrowKeySelection = true;
                    _mainWindow.KidNameComboBox.SelectedIndex -= 1;
                    e.Handled = true;
                }

                break;

            case Key.Enter:
                if (_mainWindow.KidNameComboBox.IsDropDownOpen)
                {
                    _mainWindow.KidNameComboBox.SelectedItem = _mainWindow.KidNameComboBox.Items.CurrentItem;
                    _mainWindow.KidNameComboBox.IsDropDownOpen = false;
                }

                break;
        }
    }

    public async void KidNameComboBox_Loaded()
    {
        var selectedGroup = (_mainWindow.GroupDropdown.SelectedItem as ComboBoxItem)?.Content as string ?? "Bären";
        _allKidNames = await GetKidNamesForGroupAsync(selectedGroup);
        _mainWindow.KidNameComboBox.ItemsSource = _allKidNames;

        if (_mainWindow.KidNameComboBox.Template.FindName("PART_EditableTextBox", _mainWindow.KidNameComboBox) is
            TextBox textBox)
            textBox.TextChanged += KidNameComboBoxTextBox_TextChanged;
        _mainWindow.KidNameComboBox.SelectionChanged += KidNameComboBox_SelectionChanged;
    }

    private void KidNameComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_isArrowKeySelection || sender is not ComboBox comboBox ||
            comboBox.Template.FindName("PART_EditableTextBox", comboBox) is not TextBox textBox) return;

        var currentText = textBox.Text;
        comboBox.SelectedItem = e.AddedItems[0];
        textBox.Text = currentText;
        textBox.CaretIndex = textBox.Text.Length;
    }

    private void KidNameComboBoxTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (sender is not TextBox textBox) return;

        var currentText = textBox.Text;

        if (string.IsNullOrEmpty(currentText))
        {
            _mainWindow.KidNameComboBox.ItemsSource = _allKidNames;
        }
        else
        {
            var filteredNames = _allKidNames
                .Where(name => name.StartsWith(currentText, StringComparison.OrdinalIgnoreCase)).ToList();
            _mainWindow.KidNameComboBox.ItemsSource = filteredNames;
        }

        textBox.Text = currentText;
        textBox.CaretIndex = currentText.Length;

        if (!_mainWindow.KidNameComboBox.IsDropDownOpen && _mainWindow.KidNameComboBox.HasItems)
            _mainWindow.KidNameComboBox.IsDropDownOpen = true;
    }

    public async void OnGroupSelected(SelectionChangedEventArgs e)
    {
        if (_mainWindow.KidNameComboBox == null) return;
        if (string.IsNullOrEmpty(_mainWindow.HomeFolder))
        {
            var result = ShowMessage("Möchten Sie das Hauptverzeichnis ändern?", MessageType.Info,
                "Hauptverzeichnis nicht festgelegt", MessageBoxButton.YesNo);

            if (result == MessageBoxResult.Yes)
            {
                using var dialog = new FolderBrowserDialog();
                var dialogResult = dialog.ShowDialog();
                if (dialogResult == DialogResult.OK)
                    _mainWindow.SetHomeFolder(dialog.SelectedPath);
                else
                    return;
            }
            else
            {
                return;
            }
        }

        if (e.AddedItems.Count > 0 && e.AddedItems[0] is ComboBoxItem { Content: string selectedGroup } &&
            !string.IsNullOrEmpty(selectedGroup))
        {
            _allKidNames = await GetKidNamesForGroupAsync(selectedGroup);
            KidNameComboBox_Loaded();
        }
        else
        {
            LogMessage("No valid group selected.", LogLevel.Warning);
        }
    }

    public async Task<List<string>> GetKidNamesForGroupAsync(string groupName)
    {
        var path = groupName switch
        {
            "Bären" => @"Entwicklungsberichte\Baeren Entwicklungsberichte\Aktuell",
            "Löwen" => @"Entwicklungsberichte\Loewen Entwicklungsberichte\Aktuell",
            "Schnecken" => @"Entwicklungsberichte\Schnecken Entwicklungsberichte\Aktuell",
            _ => string.Empty
        };
        return await GetKidNamesFromDirectoryAsync(path);
    }

    private async Task<List<string>> GetKidNamesFromDirectoryAsync(string groupPath)
    {
        if (_mainWindow.HomeFolder != null)
        {
            var fullPath = Path.Combine(_mainWindow.HomeFolder, groupPath);
            if (!Directory.Exists(fullPath)) return new List<string>();

            var directories = await Task.Run(() => Directory.GetDirectories(fullPath));
            return directories.Select(Path.GetFileName).OfType<string>().ToList();
        }

        LogMessage("_homeFolder is not set.", LogLevel.Warning);
        return new List<string>();
    }
}