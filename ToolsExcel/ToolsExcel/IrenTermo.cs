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
    public partial class IrenTermo
    {
        public const string CATEGORIA = "IREN_60T";

        #region Codice generato dalla finestra di progettazione di VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(this.IrenTermo_Startup);
            this.Shutdown += new System.EventHandler(this.IrenTermo_Shutdown);
            this.Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(this.IrenTermo_Change);

        }

        #endregion

        private void IrenTermo_Shutdown(object sender, EventArgs e)
        {

        }

        private void IrenTermo_Startup(object sender, EventArgs e)
        {
            Sheet<IrenTermo> s = new Sheet<IrenTermo>(this);
            s.Clear();
            s.LoadStructure();
        }

        public void UpdateStructure()
        {
            
        }

        private void IrenTermo_Change(Excel.Range Target)
        {
            
        }

    }
}
