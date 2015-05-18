using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using System.Configuration;
using System.IO;
using System.Deployment.Application;
using System.Reflection;
using System.Globalization;
using Iren.ToolsExcel.Utility;
using Iren.ToolsExcel.Base;

// ************************************************************* OFFERTE MGP ************************************************************* //

namespace Iren.ToolsExcel
{
    public partial class ThisWorkbook
    {
        #region Variabili

        public System.Version Version 
        { 
            get 
            {
                try
                {
                    return ApplicationDeployment.CurrentDeployment.CurrentVersion;
                }
                catch (Exception)
                {
                    return Assembly.GetExecutingAssembly().GetName().Version;
                }
            }
        }

        #endregion

        #region Codice generato dalla finestra di progettazione di VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.BeforeClose += new Microsoft.Office.Interop.Excel.WorkbookEvents_BeforeCloseEventHandler(this.ThisWorkbook_BeforeClose);
            this.SheetSelectionChange += new Microsoft.Office.Interop.Excel.WorkbookEvents_SheetSelectionChangeEventHandler(Handler.CellClick);
            this.WindowActivate += new Microsoft.Office.Interop.Excel.WorkbookEvents_WindowActivateEventHandler(this.ThisWorkbook_WindowActivate);
            this.Startup += new System.EventHandler(this.ThisWorkbook_Startup);
        }

        #endregion

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            ThisApplication.ScreenUpdating = false;
            ThisApplication.Iteration = true;
            ThisApplication.MaxIterations = 100;

            foreach (Excel.Worksheet ws in Sheets)
            {
                ws.Activate();
                ws.Range["A1"].Select();
                Application.ActiveWindow.ScrollRow = 1;
                if (ws.Name != "Main")
                    Application.ActiveWindow.ScrollColumn = 1;
            }

            Globals.Main.Select();
            Globals.ThisWorkbook.Application.WindowState = Excel.XlWindowState.xlMaximized;

            DateTime dataAttiva = DateTime.ParseExact(Utilities.AppSettings("DataInizio"), "yyyyMMdd", CultureInfo.InvariantCulture);
            bool emergenza = Utilities.Init(Utilities.AppSettings("DB"), Utilities.AppSettings("AppID"), dataAttiva, Globals.ThisWorkbook.Base, Version);

            Sheet.Proteggi(false);

            Riepilogo r = new Riepilogo(this.Sheets["Main"]);

            if (emergenza)
                r.RiepilogoInEmergenza();

            r.InitLabels();

            Style.StdStyles();
            Utility.Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogAccesso, "Log on - " + Environment.UserName + " - " + Environment.MachineName);

            Sheet.Proteggi(true);
            ThisApplication.ScreenUpdating = true;
        }
        private void ThisWorkbook_BeforeClose(ref bool Cancel)
        {
            ThisApplication.ScreenUpdating = false;
            if (Globals.ThisWorkbook.ThisApplication.DisplayDocumentActionTaskPane)
                Globals.ThisWorkbook.ThisApplication.DisplayDocumentActionTaskPane = false;

            Globals.Main.Select();
            Globals.ThisWorkbook.Application.WindowState = Excel.XlWindowState.xlMaximized;

            if (Simboli.ModificaDati)
            {
                Application.ScreenUpdating = false;
                Sheet.Proteggi(false);
                Simboli.ModificaDati = false;
                Sheet.AbilitaModifica(false);
                Sheet.SalvaModifiche();
                Sheet.Proteggi(true);
                Application.ScreenUpdating = true;
            }
            DataBase.SalvaModificheDB();
            this.Save();
        }
        private void ThisWorkbook_WindowActivate(Excel.Window Wn)
        {
            try
            {
                Globals.Ribbons.ToolsExcelRibbon.RibbonUI.ActivateTab(Globals.Ribbons.ToolsExcelRibbon.FrontOffice.ControlId.CustomId);
            }
            catch (Exception)
            {

            }
        }
    }
}
