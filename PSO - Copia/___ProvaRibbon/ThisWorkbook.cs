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
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace ProvaRibbon
{
    public partial class ThisWorkbook
    {
        [CachedAttribute()]
        public DataSet dsRibbonLayout = new DataSet("RibbonLayout");


        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            //DataTable dt = new DataTable("Ciccio")
            //{
            //    Columns = 
            //    {
            //        {"Bubu", typeof(string)},
            //        {"Bubu2", typeof(string)}
            //    }
            //};

            //DataRow r = dt.NewRow();
            //r["Bubu"] = "Mannaggia";
            //r["Bubu2"] = "Mannaggia2";
            //dt.Rows.Add(r);



            //dsRibbonLayout.Tables.Add(dt);
            //StringWriter strWriter = new StringWriter();
            //XmlWriter xmlWriter = XmlWriter.Create(strWriter);


            //dt.Namespace = "CiccioBenzina";
            //dt.WriteXml(xmlWriter);

            //XElement root = XElement.Parse(strWriter.ToString();            

            ////XElement root = XElement.Parse(strWriter.ToString());
            ////XNamespace ns = "CiccioBenzina";

            //Microsoft.Office.Core.CustomXMLPart part;

            //try { CustomXMLParts["CiccioBenzina"].Delete(); }
            //catch { }

            //part = CustomXMLParts.Add(XML: root.ToString(SaveOptions.DisableFormatting));
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
