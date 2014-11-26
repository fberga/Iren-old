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
using Iren.FrontOffice.Core;
using System.Configuration;
using System.IO;

namespace Iren.FrontOffice.Tools
{
    public partial class ThisWorkbook
    {
        #region Variabili

        public struct Parameters
        {
            public const int DATA_ORE_TOT = 24;
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
            this.Startup += new System.EventHandler(this.ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(this.ThisWorkbook_Shutdown);

        }

        #endregion

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            CommonFunctions.Init(ConfigurationManager.AppSettings["DB"], CommonFunctions.AppIDs.SISTEMA_COMANDI, DateTime.Now);

            Globals.Main.Select();
            Globals.ThisWorkbook.Application.WindowState = Excel.XlWindowState.xlMaximized;

            CommonFunctions.AggiornaStrutturaDati();

            //TODO riabilitare log!!
            //CommonFunctions.DB.InsertLog(DataBase.TipologiaLOG.LogAccesso, "Log on - " + Environment.UserName + " - " + Environment.MachineName);
        }

        private void ThisWorkbook_BeforeClose(ref bool Cancel)
        {
            CommonFunctions.Close();
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
            
        }
    }
}
