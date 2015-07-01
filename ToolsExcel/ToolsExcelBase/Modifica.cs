using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Base
{
    public abstract class AModifica
    {
        public abstract void Range(object Sh, Excel.Range Target);
    }

    public class Modifica : AModifica
    {
        public override void Range(object Sh, Excel.Range Target)
        {
            return;
        }
    }
}
