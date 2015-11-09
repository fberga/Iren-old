using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace Iren.ToolsExcel.UserConfig
{
    public class UserConfiguration : ConfigurationSection
    {
        [ConfigurationProperty("", IsDefaultCollection = true)]
        public UserConfigCollection Items
        {
            get { return (UserConfigCollection)base[""]; }
            set { base[""] = value; }
        }
    }
}
