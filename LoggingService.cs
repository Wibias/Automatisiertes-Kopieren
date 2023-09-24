using Serilog;
using System;

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
            _mainWindow.ShowError(userMessage);
        }

        public void HandleError(string message)
        {
            _mainWindow.ShowError(message);
        }
    }
}
