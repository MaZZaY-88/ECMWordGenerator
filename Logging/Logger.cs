using System;
using System.IO;

namespace ECMWordGenerator.Logging
{
    public static class Logger
    {
        private static string GetLogFilePath()
        {
            string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString());
            Directory.CreateDirectory(logDirectory);
            return Path.Combine(logDirectory, $"{DateTime.Now.Day}.log");
        }

        public static void Log(string message, bool isError = false)
        {
            string logFilePath = GetLogFilePath();
            string logMessage = $"{DateTime.Now}: {(isError ? "Error" : "Ok")}: {message}";
            File.AppendAllText(logFilePath, logMessage + Environment.NewLine);
        }
    }
}
