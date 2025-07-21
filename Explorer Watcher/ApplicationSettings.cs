using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Explorer_Watcher
{
    public class ApplicationSettings
    {
        public ApplicationSettings() { }

        // Set settings
        public void Save()
        {
            Properties.Settings.Default.Save();
        }

        // Set parameter value
        public void SetParameter(string parameter, string value)
        {
            Properties.Settings.Default[parameter] = value;
        }
        public void SetParameter(string parameter, bool value)
        {
            Properties.Settings.Default[parameter] = value;
        }
        public void SetParameter(string parameter, int value)
        {
            Properties.Settings.Default[parameter] = value;
        }

        // Get parameter value
        public T GetParameter<T>(string parameter)
        {
            return (T)Properties.Settings.Default[parameter];
        }
    }
}
