using System;
using System.IO;
using System.Reflection;
using System.Globalization;
using System.Threading;

namespace dotnet_core_popgmail 
{
    public static class Logger 
    {
        private static CultureInfo _cultureInfo = new CultureInfo("en-US");
        private static String _dateFormat = "yyyyMMdd";

        public static void WriteLog(string message)
        {
            try{
                
                string path = System.IO.Path.Combine(AssemblyDirectory, "logs");
                if (!Directory.Exists(path)){
                    Directory.CreateDirectory(path);
                }

                string filePath = System.IO.Path.Combine(path, string.Format("{0}.log", DateTime.Now.ToString(_dateFormat, _cultureInfo)));
                if (!File.Exists(filePath)){
                    using (var sw = File.CreateText(filePath))
                    {
                        sw.WriteLine(string.Format("{0} {1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", _cultureInfo), message)); 
                    }
                }
                else 
                {
                    using (var sw = File.AppendText(filePath))
                    {
                        sw.WriteLine(string.Format("{0} {1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", _cultureInfo), message));                         
                    }
                }

            }catch{
                
            }
        }

        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

    }
}