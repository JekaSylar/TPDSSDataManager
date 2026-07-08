using System;
using System.IO;

namespace TPDSSDataManager.Services
{
    public class SimpleLogger
    {
        private readonly string _logFilePath;
        private const long MaxFileSizeBytes = 1024 * 1024;

        public SimpleLogger(string targetDirectory)
        {
            _logFilePath = Path.Combine(targetDirectory, "TPDSS_Errors.txt");

            // Видаляємо старий лог при кожному новому запуску операції
            if (File.Exists(_logFilePath))
            {
                try { File.Delete(_logFilePath); } catch { }
            }
        }

        public void LogError(string context, string errorMessage)
        {
            try
            {
                if (File.Exists(_logFilePath))
                {
                    FileInfo fi = new FileInfo(_logFilePath);
                    if (fi.Length > MaxFileSizeBytes) File.Delete(_logFilePath);
                }

                string logEntry = $"[{DateTime.Now:dd.MM.yyyy HH:mm:ss}] [{context}] {errorMessage}{Environment.NewLine}";

                File.AppendAllText(_logFilePath, logEntry);
            }
            catch { }
        }
    }
}