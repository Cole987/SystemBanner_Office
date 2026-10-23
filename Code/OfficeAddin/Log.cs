//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.IO;

namespace SystemBannerOffice
{
    //%TEMP%\SystemBannerOffice.log: each start of the add-in in Office and every error with its stack trace,
    //like the crash log of SystemBanner, so a test on an old Office can be diagnosed from the file
    public static class Log
    {
        private const long MaxSize = 256 * 1024;
        public static readonly string FilePath = Path.Combine(Path.GetTempPath(), "SystemBannerOffice.log");

        public static void Write(string message)
        {
            try
            {
                FileInfo file = new FileInfo(FilePath);
                if (file.Exists && file.Length > MaxSize)
                {
                    file.Delete(); //start over rather than grow without end
                }
                File.AppendAllText(FilePath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " [" + System.Diagnostics.Process.GetCurrentProcess().ProcessName + "] " + message + Environment.NewLine);
            }
            catch (Exception)
            {
                //logging must never break Office
            }
        }

        public static void Error(string where, Exception ex)
        {
            Exception inner = ex is System.Reflection.TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
            Write("ERROR in " + where + ": " + inner);
        }
    }
}
