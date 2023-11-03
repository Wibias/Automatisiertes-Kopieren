using System.Windows;
using AutoUpdaterDotNET;

namespace Automatisiertes_Kopieren;

public partial class AboutWindow
{
    public AboutWindow(string? version)
    {
        InitializeComponent();

        // Set the version information in the VersionTextBlock
        VersionTextBlock.Text = version;
    }

    private void OnOkButtonClicked(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void OnSearchUpdateClicked(object sender, RoutedEventArgs e)
    {
        AutoUpdater.Start("https://raw.githubusercontent.com/enkama/Automatisiertes-Kopieren/main/autoupdater.xml");
    }
}