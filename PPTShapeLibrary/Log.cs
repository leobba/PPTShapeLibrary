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
                        // 轮转：优先删（同进程内）→ 失败则改名 .old（可能被其他进程占用时也能腾出空间）
                        try { File.Delete(path); }
                        catch
                        {
                            try
                            {
                                string old = Path.Combine(LogDir, "app.old.log");
                                if (File.Exists(old)) { try { File.Delete(old); } catch { } }
                                try { File.Move(path, old); } catch { }
                            }
                            catch { }
                        }
                    }
                    // FileShare.ReadWrite：多个 PowerPoint 实例共写同一日志不互斥崩溃
                    using (var fs = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                    using (var sw = new StreamWriter(fs))
                    {
                        sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " " + msg);
                    }
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
