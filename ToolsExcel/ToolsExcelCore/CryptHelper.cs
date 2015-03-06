using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace Iren.ToolsExcel.Core
{
    public class CryptHelper
    {
        public static void CryptSection(params string[] sections)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            string provider = "RsaProtectedConfigurationProvider";

            foreach (string sectionName in sections)
            {
                ConfigurationSection section = config.GetSection(sectionName);
                if (section != null)
                {
                    if (!section.SectionInformation.IsProtected)
                    {
                        if (!section.ElementInformation.IsLocked)
                        {
                            section.SectionInformation.ProtectSection(provider);

                            section.SectionInformation.ForceSave = true;
                            config.Save(ConfigurationSaveMode.Modified);
                        }
                    }
                }
            }
        }
    }
}
