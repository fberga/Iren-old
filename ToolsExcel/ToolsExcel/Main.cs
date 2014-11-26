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

namespace Iren.FrontOffice.Tools
{
    public partial class Main
    {
        #region Codice generato dalla finestra di progettazione di VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Main_Startup);
            this.Shutdown += new System.EventHandler(Main_Shutdown);
        }

        #endregion

        private void Main_Startup(object sender, System.EventArgs e)
        {
        }

        private void Main_Shutdown(object sender, System.EventArgs e)
        {
        }
    }
}
