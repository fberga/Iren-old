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
using ComFunc = Iren.FrontOffice.Tools.CommonFunctions;

namespace Iren.FrontOffice.Tools
{
    public partial class Log
    {
        #region Variabili

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

        //public object[,] DataTableToObjArray(System.Data.DataTable dt)
        //{
        //    object[,] o = new object[dt.Rows.Count, dt.Columns.Count];

        //    int i = 0;
        //    foreach (System.Data.DataRow row in dt.Rows)
        //    {
        //        int j = 0;
        //        foreach (System.Data.DataColumn col in dt.Columns)
        //        {
        //            o[i, j++] = row[col];
        //        }
        //        i++;
        //    }
        //    return o;
        //}

        #endregion

        #region Callbacks

        private void Log_Startup(object sender, EventArgs e)
        {
            ListObject logObj;
            try
            {
                logObj = Globals.Factory.GetVstoObject(ListObjects["LogList"]);
            }
            catch (Exception)
            {
                logObj = Controls.AddListObject(Range["A1"], "LogList");
            }
            logObj.AutoSetDataBoundColumnHeaders = true;
            logObj.DataSource = ComFunc.LocalDB;
            logObj.DataMember = ComFunc.Tab.LOG;
            logObj.Range.EntireColumn.AutoFit();
        }

        private void Log_Shutdown(object sender, EventArgs e)
        {
        }

        #endregion

    }
}
