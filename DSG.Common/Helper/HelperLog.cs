using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DSG.Common.Helper
{
    public class HelperLog
    {
        private static void WriteLine(string path, string fileName, string msg)
        {
            if (!System.IO.Directory.Exists(path))
                System.IO.Directory.CreateDirectory(path);
            string fileFullName = Path.Combine(path, fileName);
            using (FileStream stream = new FileStream(fileFullName, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
            {
                StreamWriter write = new StreamWriter(stream);
                write.WriteLine(DateTime.Now.ToString() + " " + msg);
                write.Flush();
                write.Close();
            }
        }

        public static void SensorLogWritter(string msg)
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "SensorLog");
            string fileName = DateTime.Now.ToString("yyyyMMdd") + ".txt";
            WriteLine(path, fileName, msg);
        }

        public static void LogWritter(string msg)
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "Log");
            string fileName = DateTime.Now.ToString("yyyyMMdd") + ".txt";
            WriteLine(path, fileName, msg);
        }

    }
}
