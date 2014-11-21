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
        public static DataBase _db;
        private DataSet _localDB;
        public const string NS = "Iren.FrontOffice.SistemaComandi";


        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            _db = new DataBase(ConfigurationManager.AppSettings["DB"]);

            string usr = Environment.UserName;
            DataSet ds = new DataSet("LocalDB");
            
            DataTable dtUtente = _db.Select("spUtente", new QryParams() { { "@CodUtenteWindows", usr } });
            dtUtente.TableName = "Utente";

            DataTable dtLog = _db.Select("spLog", new QryParams() { { "@IdApplicazione", "8" }, {"@Data", "20140929"}});
            dtLog.TableName = "Log";

            ds.Namespace = NS;
            ds.Prefix = "LocalDB";
            ds.Tables.Add(dtUtente);
            ds.Tables.Add(dtLog);

            StringWriter sw = new StringWriter();
            ds.WriteXml(sw);
            string result = sw.ToString();

            Office.CustomXMLPart user = Globals.ThisWorkbook.CustomXMLParts.Add(result, missing);
            user.NamespaceManager.AddNamespace("Iren.FrontOffice.Tools", "User");

            Office.CustomXMLParts sss = Globals.ThisWorkbook.CustomXMLParts.SelectByNamespace("User");

            //foreach (Office.CustomXMLPart xmlPart in Globals.ThisWorkbook.CustomXMLParts.SelectByNamespace("AAA"))
            //{
            //    if (!xmlPart.BuiltIn)
            //    {
            //        MessageBox.Show(xmlPart.DocumentElement.NamespaceURI);
            //        MessageBox.Show(xmlPart.DocumentElement.XML);
            //    }

            //}

        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Codice generato dalla finestra di progettazione di VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
