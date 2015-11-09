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

namespace ProvaRibbon
{
    public partial class Foglio3
    {
        private void Foglio3_Startup(object sender, System.EventArgs e)
        {
        }

        private void Foglio3_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Codice generato dalla finestra di progettazione di VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Foglio3_Startup);
            this.Shutdown += new System.EventHandler(Foglio3_Shutdown);
        }

        #endregion

    }
}
