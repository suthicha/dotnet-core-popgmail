using System;
using System.Reflection;
using System.IO;

namespace dotnet_core_popgmail {
    static public class Utils {
        static public string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        static public string InitPath(string name){
            string path = Path.Combine(AssemblyDirectory, name);
            if (!Directory.Exists(path)){
                Directory.CreateDirectory(path);
            }
            return path;
        }
    }
}