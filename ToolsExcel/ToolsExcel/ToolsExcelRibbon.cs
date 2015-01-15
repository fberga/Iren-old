using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Iren.FrontOffice.Base;
using System.Configuration;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using Iren.FrontOffice.Core;

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

            DataView categorie = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.CATEGORIA].DefaultView;
            categorie.RowFilter = "Operativa = 1";

            foreach (DataRowView categoria in categorie)
            {
                Excel.Worksheet ws;
                try 
                {
                    ws = Globals.ThisWorkbook.Worksheets[categoria["DesCategoria"].ToString()];
                }
                catch
                {
                    ws = (Excel.Worksheet)Globals.ThisWorkbook.Worksheets.Add(Globals.ThisWorkbook.Worksheets["Log"]);
                    ws.Name = categoria["DesCategoria"].ToString();
                    ws.Select();
                    Globals.ThisWorkbook.Application.Windows[1].DisplayGridlines = false;
                }
            }

            Riepilogo main = new Riepilogo(Globals.ThisWorkbook.Sheets["Main"]);
            main.LoadStructure();

            foreach (Excel.Worksheet ws in Globals.ThisWorkbook.Sheets) 
            {
                if(ws.Name != "Log" && ws.Name != "Main")
                {
                    Sheet s = new Sheet(ws);
                    s.LoadStructure();
                }
            }

            Globals.Main.Select();
            Globals.ThisWorkbook.Application.WindowState = Excel.XlWindowState.xlMaximized;

            //TODO riabilitare log!!
            //Globals.Log.Unprotect();
            //CommonFunctions.InsertLog(DataBase.TipologiaLOG.LogModifica, "Aggiorna struttura");
            //Globals.Log.Protect();

            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }
    }
}
