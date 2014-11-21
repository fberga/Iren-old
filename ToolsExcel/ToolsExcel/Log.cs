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

namespace Iren.FrontOffice.Tools
{
    public partial class Log
    {
        #region Variabili

        private DataSet _localDB;

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

        public object[,] DataTableToObjArray(System.Data.DataTable dt)
        {
            object[,] o = new object[dt.Rows.Count, dt.Columns.Count];

            int i = 0;
            foreach (System.Data.DataRow row in dt.Rows)
            {
                int j = 0;
                foreach (System.Data.DataColumn col in dt.Columns)
                {
                    o[i, j++] = row[col];
                }
                i++;
            }
            return o;
        }

        #endregion

        #region Callbacks

        private void Log_Startup(object sender, EventArgs e)
        {
            _localDB = new DataSet();

            foreach (Office.CustomXMLPart xmlPart in Globals.ThisWorkbook.CustomXMLParts.SelectByNamespace(ThisWorkbook.NS))
            {
                StringReader sr = new StringReader(xmlPart.XML);
                _localDB.ReadXml(sr);
            }

            DataTable dt = _localDB.Tables["Log"];

            Range[Cells[2, 1], Cells[dt.Rows.Count, dt.Columns.Count]].Value = DataTableToObjArray(dt);

        }

        private void Log_Shutdown(object sender, EventArgs e)
        {
        }

        #endregion

    }
}
