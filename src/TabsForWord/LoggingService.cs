using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace TabsForWord
{
    /// <summary>
    /// A plain file log: %LOCALAPPDATA%\TabsForWord\Logs\TabsForWord-yyyyMMdd.log.
    /// Document CONTENT is never written; technical events do contain document
    /// NAMES and PATHS, and these files end up in the diagnostic ZIP whose
    /// contents are spelled out for the user (INSTALL_EN.md, diagnose.ps1 output).
    /// A logging failure is never propagated to the caller.
    /// </summary>
    public static class LoggingService
    {
        private const int MaxLogFiles = 7;
        private static readonly object Sync = new object();
        private static string _logDir;
        private static bool _initFailed;

        public static void Info(string message) { Write("INFO ", message, null); }
        public static void Warn(string message) { Write("WARN ", message, null); }
        public static void Error(string message, Exception ex = null) { Write("ERROR", message, ex); }

        private static void Write(string level, string message, Exception ex)
        {
            try
            {
                lock (Sync)
                {
                    var dir = EnsureLogDir();
                    if (dir == null) return;

                    var sb = new StringBuilder();
                    sb.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture));
                    sb.Append(" [").Append(level).Append("] ").Append(message);
                    if (ex != null)
                    {
                        sb.AppendLine();
                        sb.Append("        ").Append(ex.GetType().FullName).Append(": ").Append(ex.Message);
                        if (ex.StackTrace != null)
                        {
                            sb.AppendLine();
                            sb.Append(ex.StackTrace);
                        }
                    }
                    sb.AppendLine();

                    var file = Path.Combine(dir, "TabsForWord-" + DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + ".log");
                    File.AppendAllText(file, sb.ToString(), Encoding.UTF8);
                }
            }
            catch
            {
                // Logging must never bring the add-in down.
            }
        }

        private static string EnsureLogDir()
        {
            if (_initFailed) return null;
            if (_logDir != null) return _logDir;
            try
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "TabsForWord", "Logs");
                Directory.CreateDirectory(dir);
                CleanupOldLogs(dir);
                _logDir = dir;
                return dir;
            }
            catch
            {
                _initFailed = true;
                return null;
            }
        }

        private static void CleanupOldLogs(string dir)
        {
            try
            {
                var old = new DirectoryInfo(dir).GetFiles("TabsForWord-*.log")
                    .OrderByDescending(f => f.Name)
                    .Skip(MaxLogFiles)
                    .ToList();
                foreach (var f in old)
                {
                    try { f.Delete(); } catch { }
                }
            }
            catch { }
        }
    }
}
