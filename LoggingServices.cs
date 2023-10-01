using Serilog;
using System;
using System.IO;
using System.Windows;

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
        public void LogAndShowMessage(string logMessage, string userMessage, LogLevel logLevel = LogLevel.Info, MessageType messageType = MessageType.Info, string? title = null)
        {
            LogMessage(logMessage, logLevel);
            ShowMessage(userMessage, messageType, title);
        }

        public enum LogLevel
        {
            Error,
            Info,
            Warning
        }

        public void LogMessage(string message, LogLevel level = LogLevel.Info, Exception? exception = null)
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
                case LogLevel.Info:
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
            Info,
            Warning
        }

        public MessageBoxResult ShowMessage(string message, MessageType type = MessageType.Info, string? title = null, MessageBoxButton button = MessageBoxButton.OK)
        {
            MessageBoxImage icon;

            switch (type)
            {
                case MessageType.Error:
                    icon = MessageBoxImage.Error;
                    title = title ?? "Fehler";
                    break;
                case MessageType.Info:
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
