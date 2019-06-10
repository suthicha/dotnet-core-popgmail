using System;
using System.IO;
using Newtonsoft.Json;

namespace dotnet_core_popgmail {
    public class MailSettings {
        public String Host {get;set;}
        public Int32 Port {get;set;}
        public Boolean isSSL {get;set;}
        public String EmailAddress {get;set;}
        public String EmailPassword {get;set;}

        public MailSettings Read(){
            try{
                string configPath = Path.Combine(Utils.AssemblyDirectory, "config.json");
                string jsonContent = File.ReadAllText(configPath);
                return JsonConvert.DeserializeObject<MailSettings>(jsonContent);
            }catch{}
            return null;
        }
    }


}