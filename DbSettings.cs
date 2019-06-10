using System;
using System.IO;
using Newtonsoft.Json;

namespace dotnet_core_popgmail {
    public class DbSettings {

        public string DB_Server {get;set;}
        public string DB_UserId {get;set;}
        public string DB_Password {get;set;}
        public string DB_Database {get;set;}
        
        public string GetConnectionString(){
            try{
                string configPath = Path.Combine(Utils.AssemblyDirectory, "config.json");
                string jsonContent = File.ReadAllText(configPath);
                var obj =  JsonConvert.DeserializeObject<DbSettings>(jsonContent);
                return String.Format("data source={0};uid={1};password={2};database={3}", obj.DB_Server, obj.DB_UserId, obj.DB_Password, obj.DB_Database);
            }catch{}
            return string.Empty;
        }
    }
}