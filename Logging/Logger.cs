using ECMWordGenerator.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ECMWordGenerator.Logging
{
    public static class Logger
    {
        private static readonly string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");

        static Logger()
        {
            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
            }
        }

        public static void Log(string message, bool isError = false)
        {
            string logFilePath = GetLogFilePath();
            string logMessage = $"{DateTime.Now:dd.MM.yyyy HH:mm:ss}: {(isError ? "Error" : "Ok")}: {message}";

            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                writer.WriteLine(logMessage);
            }
        }

        private static string GetLogFilePath()
        {
            string yearDirectory = Path.Combine(logDirectory, DateTime.Now.ToString("yyyy"));
            string monthDirectory = Path.Combine(yearDirectory, DateTime.Now.ToString("MM"));
            string logFilePath = Path.Combine(monthDirectory, $"{DateTime.Now:dd.MM.yyyy}.log");

            if (!Directory.Exists(yearDirectory))
            {
                Directory.CreateDirectory(yearDirectory);
            }

            if (!Directory.Exists(monthDirectory))
            {
                Directory.CreateDirectory(monthDirectory);
            }

            return logFilePath;
        }

        /// <summary>
        /// Formats the Data attribute values for logging.
        /// </summary>
        /// <param name="data">The list of key-value pairs to format.</param>
        /// <returns>A string representation of the Data attribute values.</returns>
        public static string FormatData(List<Item> data)
        {
            return string.Join(", ", data.Select(d => $"{{Placeholder: {d.Placeholder}, Value: {d.Value}}}"));
        }
    }
}
