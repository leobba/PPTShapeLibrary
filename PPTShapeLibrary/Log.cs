using System;
using System.IO;

namespace PPTShapeLibrary
{
    /// <summary>简单文件日志（E3）：%APPDATA%\PPTShapeLibrary\logs\app.log，超过 1MB 自动轮转。</summary>
    public static class Log
    {
        private static readonly object LockObj = new object();
        private static string _dir;

        private static string LogDir
        {
            get
            {
                if (_dir == null)
                {
                    _dir = Path.Combine(ShapeLibraryService.RootDir, "logs");
                    try { Directory.CreateDirectory(_dir); } catch { }
                }
                return _dir;
            }
        }

        public static void Write(string msg)
        {
            try
            {
                lock (LockObj)
                {
                    string path = Path.Combine(LogDir, "app.log");
                    if (File.Exists(path) && new FileInfo(path).Length > 1024 * 1024)
                    {
                        try { File.Delete(path); } catch { }
                    }
                    File.AppendAllText(path,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " " + msg + Environment.NewLine);
                }
            }
            catch { }
        }

        public static void Error(Exception ex, string context)
        {
            Write("ERROR [" + context + "]: " + (ex == null ? "null" : ex.ToString()));
        }
    }
}
