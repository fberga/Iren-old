using System;
using System.IO;
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
using Iren.ToolsExcel.Base;

namespace Iren.ToolsExcel
{
    public partial class Log
    {
        #region Variabili

        public ListObject _logObj;

        #endregion

        #region Codice generato dalla finestra di progettazione di VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(this.Log_Startup);
            this.Shutdown += new System.EventHandler(this.Log_Shutdown);

        }

        #endregion

        #region Metodi

        #endregion

        #region Callbacks

        private void Log_Startup(object sender, EventArgs e)
        {            
        }

        private void Log_Shutdown(object sender, EventArgs e)
        {
        }

        #endregion

    }
}
