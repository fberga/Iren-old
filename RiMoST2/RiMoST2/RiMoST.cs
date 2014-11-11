using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace RiMoST2
{
    public partial class RiMoST
    {
        private void RiMoST_Load(object sender, RibbonUIEventArgs e)
        {
            IList<IRibbonExtension> ribbonList = Globals.Ribbons.Base;

            foreach (IRibbonExtension r in ribbonList)
            {

                r.ToString();
            }
        }
    }
}
