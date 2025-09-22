using System;
using System.IO;
using System.IO.Compression;

namespace Guifender
{
    public static class Logger
    {
        private static readonly string _logFilePath;
        private static readonly string _logDirectory;
        private static readonly object _lock = new object();

        static Logger()
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _logDirectory = Path.Combine(appDataPath, "Guifender");
            Directory.CreateDirectory(_logDirectory);
            _logFilePath = Path.Combine(_logDirectory, "Guifender.log");
        }

        public static void ArchiveLogFileIfNeeded()
        {
            lock (_lock)
            {
                if (!File.Exists(_logFilePath)) { return; }

                var lastWriteTime = File.GetLastWriteTime(_logFilePath);
                var now = DateTime.Now;

                if (lastWriteTime.Year < now.Year || lastWriteTime.Month < now.Month)
                {
                    try
                    {
                        string archiveFileName = $"Guifender_{lastWriteTime:yyyyMM}.log";
                        string archiveFilePath = Path.Combine(_logDirectory, archiveFileName);

                        // Rename the current log file
                        File.Move(_logFilePath, archiveFilePath);

                        // Create a zip archive
                        string zipFilePath = Path.ChangeExtension(archiveFilePath, ".zip");
                        if (File.Exists(zipFilePath)) { File.Delete(zipFilePath); } // Avoid exception if zip already exists
                        
                        using (var zipArchive = ZipFile.Open(zipFilePath, ZipArchiveMode.Create))
                        {
                            zipArchive.CreateEntryFromFile(archiveFilePath, archiveFileName);
                        }

                        // Delete the temporary log file
                        File.Delete(archiveFilePath);
                    }
                    catch (Exception ex)
                    {
                        // Log archival failure to the new log file
                        Write($"Failed to archive log file: {ex.Message}");
                    }
                }
            }
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