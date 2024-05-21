using Serilog;
using System;
using System.IO;
using System.Windows;

namespace Automatisiertes_Kopieren.Helper;

public static class LoggingHelper
{
    public enum LogLevel
    {
        Error,
        Info,
        Warning,
        Debug
    }

    public enum MessageType
    {
        Error,
        Info,
        Warning
    }

    public static void InitializeLogger()
    {
        var logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Automatisiertes_Kopieren", "logs");
        var logFilePath = Path.Combine(logDirectory, "log-.txt");
        Log.Logger = new LoggerConfiguration()
            .WriteTo.File(logFilePath, rollingInterval: RollingInterval.Day, retainedFileCountLimit: 10)
            .CreateLogger();
    }

    public static void LogAndShowMessage(string logMessage, string? userMessage, LogLevel logLevel = LogLevel.Info,
        MessageType messageType = MessageType.Info, string? title = null)
    {
        LogMessage(logMessage, logLevel);
        ShowMessage(userMessage, messageType, title);
    }

    public static void LogException(Exception ex, string message, LogLevel logLevel = LogLevel.Error)
    {
        LogMessage(message, logLevel, ex);

        ShowMessage(message, MessageType.Error);
    }

    public static void LogMessage(string message, LogLevel level = LogLevel.Info, Exception? exception = null)
    {
        switch (level)
        {
            case LogLevel.Error:
                if (exception != null)
                    Log.Error(exception, message);
                else
                    Log.Error(message);
                break;
            case LogLevel.Info:
                Log.Information(message);
                break;
            case LogLevel.Warning:
                Log.Warning(message);
                break;
            case LogLevel.Debug:
                Log.Debug(message);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(level), level, null);
        }
    }

    public static MessageBoxResult ShowMessage(string? message, MessageType type = MessageType.Info,
        string? title = null,
        MessageBoxButton button = MessageBoxButton.OK)
    {
        MessageBoxImage icon;

        switch (type)
        {
            case MessageType.Error:
                icon = MessageBoxImage.Error;
                title ??= "Fehler";
                break;
            case MessageType.Info:
                icon = MessageBoxImage.Information;
                title ??= "Information";
                break;
            case MessageType.Warning:
            default:
                icon = MessageBoxImage.None;
                break;
        }

        return MessageBox.Show(message, title, button, icon);
    }
}