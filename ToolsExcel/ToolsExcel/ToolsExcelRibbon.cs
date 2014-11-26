using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace Iren.FrontOffice.Tools
{
    public partial class ToolsExcelRibbon
    {
        private void ToolsExcelRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnAggiornaStruttura_Click(object sender, RibbonControlEventArgs e)
        {
            CommonFunctions.AggiornaStrutturaDati();
        }
    }
}
