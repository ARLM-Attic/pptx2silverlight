using System;
using System.Configuration;

namespace HCLT.MSFT.TIL.PPT2007Convertor
{
    public static class ConfigManager
    {
        private static string GetValue(string key)
        {
            return ConfigurationSettings.AppSettings[key];
        }

        public static string SilverLightTemplatePath
        {
            get
            {
                return ConfigurationSettings.AppSettings["SilverLightTemplatePath"];
            }
        }
    }
}
