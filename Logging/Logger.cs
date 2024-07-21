using System;
using System.IO;

namespace ECMWordGenerator.Logging
{
    public static class Logger
    {
        /// <summary>
        /// Gets the log file path based on the current date.
        /// </summary>
        /// <returns>The path to the log file for the current date.</returns>
        private static string GetLogFilePath()
        {
            string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString());
            Directory.CreateDirectory(logDirectory);
            return Path.Combine(logDirectory, $"{DateTime.Now:dd.MM.yyyy}.log");
        }

        /// <summary>
        /// Logs a message to the log file.
        /// </summary>
        /// <param name="message">The message to log.</param>
        /// <param name="isError">Indicates whether the message is an error message.</param>
        public static void Log(string message, bool isError = false)
        {
            string logFilePath = GetLogFilePath();
            string logMessage = $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}: {(isError ? "Error" : "Ok")}: {message}";
            File.AppendAllText(logFilePath, logMessage + Environment.NewLine);
        }
    }
}
