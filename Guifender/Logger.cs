using System;
using System.IO;

namespace Guifender
{
    public static class Logger
    {
        private static readonly string _logFilePath;
        private static readonly object _lock = new object();

        static Logger()
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logDir = Path.Combine(appDataPath, "Guifender");
            Directory.CreateDirectory(logDir);
            _logFilePath = Path.Combine(logDir, "Guifender.log");
        }

        public static void Write(string message)
        {
            try
            {
                lock (_lock)
                {
                    string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}";
                    File.AppendAllText(_logFilePath, logMessage);
                }
            }
            catch (Exception)
            {
                // Silently fail to avoid the logger causing application crashes.
            }
        }
    }
}
