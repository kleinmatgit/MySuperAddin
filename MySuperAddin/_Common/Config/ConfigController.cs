using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MySuperAddin._Common.Config
{
    internal class ConfigController
    {
        //Retrieves the value of a config parameter; if it does not exist, returns defaultValue
        public static string GetConfigEntry(string key, string defaultValue = "")
        {
            LazyInitialize();
            if (_config.ContainsKey(key))
                return _config[key];
            return defaultValue;
        }

        //Looks up and returns the config entry for the specific key only if inputValue is empty
        public static string GetConfigEntryAsDefault(string key, string inputValue)
        {
            if (inputValue.Length == 0)
                return GetConfigEntry(key);
            else
                return inputValue;
        }

        //Overrides the value of a config parameter, for the duration of the current session
        public static bool SetConfigEntry(string key, string value)
        {
            LazyInitialize();
            _config[key] = value;
            return true;
        }

        
        #region Private members

        //Resets all config entries, reloading from the source config documents when applicable.
        //All overrides are lost.
        private static bool Reload()
        {
            _config.Clear();
            //TO DO: move to a proper config file

            //for instance:
            //_config["Namespace1.Prefix"] = "gmt";
            //_config["Namespace1.URI"] = "http://gmt/Schema/TradeCommon/1";
            //_config["Namespace2.Prefix"] = "debt";
            //_config["Namespace2.URI"] = "static.debt";

            return true;
        }

        private static void LazyInitialize()
        {
            if(!_initialised)
            {
                Reload();
                _initialised = true;
            }
        }

        private static Dictionary<string, string> _config = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
        private static bool _initialised = false;

        #endregion
    }
}
