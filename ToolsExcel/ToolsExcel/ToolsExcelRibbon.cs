using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Iren.FrontOffice.Base;
using System.Configuration;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.FrontOffice.Tools
{
    public partial class ToolsExcelRibbon
    {
        private void ToolsExcelRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnAggiornaStruttura_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            CommonFunctions.AggiornaStrutturaDati();

            Globals.IrenIdro.LoadStructure();
            Globals.IrenTermo.LoadStructure();
            Globals.Main.LoadStructure();

            Globals.Main.Select();
            Globals.ThisWorkbook.Application.WindowState = Excel.XlWindowState.xlMaximized;

            //TODO riabilitare log!!
            //CommonFunctions.DB.InsertLog(DataBase.TipologiaLOG.LogAccesso, "Log on - " + Environment.UserName + " - " + Environment.MachineName);
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }
    }
}
