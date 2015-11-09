using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace Iren.PSO.UserConfig
{
    public class UserConfigCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new UserConfigElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((UserConfigElement)element).Key;
        }

        public new UserConfigElement this[string key]
        {
            get { return (UserConfigElement)BaseGet(key); }
            set 
            {
                if (key != null && BaseGet(key) != null)
                    BaseRemoveAt(BaseIndexOf(BaseGet(key)));
                
                BaseAdd(value); 
            }
        }
    }
}
