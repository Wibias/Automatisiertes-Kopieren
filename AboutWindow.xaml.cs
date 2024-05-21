using AutoUpdaterDotNET;
using System;
using System.Net;
using System.Windows;
using System.Windows.Threading;
using static Automatisiertes_Kopieren.Helper.LoggingHelper;

namespace Automatisiertes_Kopieren;

public partial class AboutWindow
{
    private DispatcherTimer? _updateCheckTimer;

    public AboutWindow(string? version)
    {
        InitializeComponent();

        VersionTextBlock.Text = version;

        SetupAutoUpdate();
    }

    private void SetupAutoUpdate()
    {
        _updateCheckTimer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(2) };
        _updateCheckTimer.Tick += CheckForUpdates;
        _updateCheckTimer.Start();

        AutoUpdater.CheckForUpdateEvent += AutoUpdaterOnCheckForUpdateEvent;
    }

    private static void CheckForUpdates(object? sender, EventArgs e)
    {
        AutoUpdater.Start("https://raw.githubusercontent.com/enkama/Automatisiertes-Kopieren/main/autoupdater.xml");
    }

    private static void AutoUpdaterOnCheckForUpdateEvent(UpdateInfoEventArgs args)
    {
        switch (args.Error)
        {
            case null:
                {
                    if (args.IsUpdateAvailable)
                    {
                        MessageBoxResult dialogResult;
                        if (args.Mandatory.Value)
                            dialogResult = ShowMessage(
                                $@"Es ist eine neue Version {args.CurrentVersion} verfügbar. Sie verwenden die Version {args.InstalledVersion}. Dies ist ein erforderliches Update. Drücken Sie OK, um mit der Aktualisierung der Anwendung zu beginnen.",
                                MessageType.Info,
                                "Update verfügbar");
                        else
                            dialogResult = ShowMessage(
                                $@"Es ist eine neue Version {args.CurrentVersion} verfügbar. Sie verwenden die Version {args.InstalledVersion}. Möchten Sie die Anwendung jetzt aktualisieren?",
                                MessageType.Info,
                                "Update verfügbar",
                                MessageBoxButton.YesNo);

                        if (dialogResult != MessageBoxResult.Yes && dialogResult != MessageBoxResult.OK) return;
                        try
                        {
                            if (AutoUpdater.DownloadUpdate(args)) Application.Current.Shutdown();
                        }
                        catch (Exception exception)
                        {
                            LogAndShowMessage(exception.Message, exception.GetType().ToString(),
                                LogLevel.Error, MessageType.Error);
                        }
                    }

                    break;
                }
            case WebException:
                ShowMessage(
                    "Es besteht ein Problem beim Erreichen des Update-Servers. Bitte überprüfen Sie Ihre Internetverbindung und versuchen Sie es später erneut.",
                    MessageType.Error,
                    "Update-Überprüfung fehlgeschlagen");
                break;
            default:
                ShowMessage(args.Error.Message, MessageType.Error,
                    args.Error.GetType().ToString());
                break;
        }
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