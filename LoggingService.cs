using Serilog;
using System;
using System.Windows;

namespace Automatisiertes_Kopieren
{
    public class LoggingService
    {
        private readonly MainWindow _mainWindow;

        public LoggingService(MainWindow mainWindow)
        {
            _mainWindow = mainWindow ?? throw new ArgumentNullException(nameof(mainWindow));
        }

        public void LogAndShowError(string logMessage, string userMessage)
        {
            Log.Error(logMessage);
            ShowError(userMessage);
        }
        public void LogAndShowInformation(string logMessage, string userMessage)
        {
            Log.Information(logMessage);
            ShowInformation(userMessage);
        }

        public void ShowError(string message)
        {
            MessageBox.Show(message, "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        public void ShowInformation(string message, string title = "Information")
        {
            MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
