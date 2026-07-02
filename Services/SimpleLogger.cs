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

            // Удаляем старый лог при каждом новом запуске операции, 
            // чтобы пользователь не видел ошибки с прошлых попыток
            if (File.Exists(_logFilePath))
            {
                try { File.Delete(_logFilePath); } catch { }
            }
        }

        public void LogError(string context, string errorMessage)
        {
            try
            {
                // Если файл уже есть и он огромный, сбрасываем его
                if (File.Exists(_logFilePath))
                {
                    FileInfo fi = new FileInfo(_logFilePath);
                    if (fi.Length > MaxFileSizeBytes) File.Delete(_logFilePath);
                }

                string logEntry = $"[{DateTime.Now:dd.MM.yyyy HH:mm:ss}] [{context}] {errorMessage}{Environment.NewLine}";

                // Записываем ошибку. Файл создастся только в этот момент!
                File.AppendAllText(_logFilePath, logEntry);
            }
            catch { }
        }
    }
}