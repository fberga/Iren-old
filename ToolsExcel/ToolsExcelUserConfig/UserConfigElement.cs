using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace Iren.FrontOffice.UserConfig
{
    public class UserConfigElement : ConfigurationElement
    {        
        [ConfigurationProperty("key", IsRequired = true, IsKey = true)]
        public string Key
        {
            get { return (string)base["key"]; }
            set { base["key"] = value; }
        }

        [ConfigurationProperty("desc", IsRequired = false, DefaultValue="")]
        public string Desc
        {
            get { return (string)base["desc"]; }
            set { base["desc"] = value; }
        }

        [ConfigurationProperty("value", IsRequired = true)]
        public string Value
        {
            get { return Environment.ExpandEnvironmentVariables((string)base["value"]); }
            set { base["value"] = value; }
        }

        [ConfigurationProperty("default", IsRequired = false, DefaultValue="")]
        public string Default
        {
            get { return Environment.ExpandEnvironmentVariables((string)base["default"]); }
            set { base["default"] = value; }
        }

        [ConfigurationProperty("emergenza", IsRequired = false, DefaultValue="")]
        public string Emergenza
        {
            get { return Environment.ExpandEnvironmentVariables((string)base["emergenza"]); }
            set { base["emergenza"] = value; }
        }

        [ConfigurationProperty("archivio", IsRequired = false, DefaultValue = "")]
        public string Archivio
        {
            get { return Environment.ExpandEnvironmentVariables((string)base["archivio"]); }
            set { base["archivio"] = value; }
        }

        [ConfigurationProperty("visibile", IsRequired = false, DefaultValue="true")]
        public string Visibile
        {
            get { return (string)base["visibile"]; }
            set { base["visibile"] = value; }
        }

    }
}
