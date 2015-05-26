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

// ************************************************************* PROGRAMMAZIONE ************************************************************* //

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
            this.Startup += new System.EventHandler(this.ThisWorkbook_Startup);
        }

        #endregion

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            Utility.Workbook.StartUp(Base, Version);
        }
        private void ThisWorkbook_BeforeClose(ref bool Cancel)
        {
            Utility.Workbook.Close();
        }
    }
}
