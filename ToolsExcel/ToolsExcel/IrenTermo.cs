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
using System.Configuration;
using System.Globalization;
using Iren.FrontOffice.Base;

namespace Iren.FrontOffice.Tools
{
    public partial class IrenTermo
    {
        public Dictionary<string, object> config = new Dictionary<string, object>();

        #region Codice generato dalla finestra di progettazione di VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(this.IrenTermo_Startup);
            this.Shutdown += new System.EventHandler(this.IrenTermo_Shutdown);

        }

        #endregion

        private void IrenTermo_Shutdown(object sender, EventArgs e)
        {

        }

        private void IrenTermo_Startup(object sender, EventArgs e)
        {
            //inizializzo parametri da file di configurazione
            config.Add("SiglaCategoria", "IREN_60T");
            config.Add("DataInizio", DateTime.ParseExact(ConfigurationManager.AppSettings["DataInizio"], 
                "yyyyMMdd", CultureInfo.InvariantCulture));

            //Sheet<IrenTermo> s = new Sheet<IrenTermo>(this);
            //s.LoadStructure();
        }
    }
}
