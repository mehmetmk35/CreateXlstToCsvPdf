using Microsoft.Extensions.Configuration;

namespace WinFormsApp1
{
    static class Configuration
    {
        static public string Path
        {
            get
            {

                var path = Environment.CurrentDirectory;
                ConfigurationManager configurationManager = new();
                configurationManager.SetBasePath(path); ;
                configurationManager.AddJsonFile("appsettings.json");
                return configurationManager.GetSection("appSettings")["Path"];
            }
        }
    }
}
