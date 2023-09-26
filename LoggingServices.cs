using Serilog;
using System;
using System.Windows;
using System.IO;

namespace Automatisiertes_Kopieren
{
    public class LoggingService
    {

        public void InitializeLogger()
        {
            string logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Automatisiertes_Kopieren", "logs");
            string logFilePath = Path.Combine(logDirectory, "log-.txt");
            Log.Logger = new LoggerConfiguration()
                .WriteTo.File(logFilePath, rollingInterval: RollingInterval.Day, retainedFileCountLimit: 10)
                .CreateLogger();
        }
        public void LogAndShowMessage(string logMessage, string userMessage, LogLevel logLevel = LogLevel.Information, MessageType messageType = MessageType.Information, string? title = null)
        {
            LogMessage(logMessage, logLevel);
            ShowMessage(userMessage, messageType, title);
        }

        public enum LogLevel
        {
            Error,
            Information,
            Warning
        }

        public void LogMessage(string message, LogLevel level = LogLevel.Information, Exception? exception = null)
        {
            switch (level)
            {
                case LogLevel.Error:
                    if (exception != null)
                    {
                        Log.Error(exception, message);
                    }
                    else
                    {
                        Log.Error(message);
                    }
                    break;
                case LogLevel.Information:
                    Log.Information(message);
                    break;
                case LogLevel.Warning:
                    Log.Warning(message);
                    break;
                default:
                    break;
            }
        }

        public enum MessageType
        {
            Error,
            Information,
            Warning
        }

        public MessageBoxResult ShowMessage(string message, MessageType type = MessageType.Information, string? title = null, MessageBoxButton button = MessageBoxButton.OK)
        {
            MessageBoxImage icon;

            switch (type)
            {
                case MessageType.Error:
                    icon = MessageBoxImage.Error;
                    title = title ?? "Fehler";
                    break;
                case MessageType.Information:
                    icon = MessageBoxImage.Information;
                    title = title ?? "Information";
                    break;
                default:
                    icon = MessageBoxImage.None;
                    break;
            }

            return MessageBox.Show(message, title, button, icon);
        }
    }
}
