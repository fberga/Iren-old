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
using Iren.FrontOffice.Base;

namespace Iren.FrontOffice.Tools
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
            try
            {
                _logObj = Globals.Factory.GetVstoObject(ListObjects["LogList"]);
            }
            catch (Exception)
            {
                _logObj = Controls.AddListObject(Range["A1"], "LogList");
            }
            _logObj.AutoSetDataBoundColumnHeaders = true;

            if (CommonFunctions.DB.OpenConnection())
            {
                _logObj.DataSource = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.LOG].DefaultView;
                Unprotect(Simboli.pwd);
                _logObj.Range.EntireColumn.AutoFit();
                _logObj.TableStyle = "TableStyleLight16";

                Excel.Range rng = Columns[2];
                rng.NumberFormat = "dd/MM/yyyy HH:mm:ss";
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                Protect(Simboli.pwd);

                CommonFunctions.DB.CloseConnection();
            }
        }

        private void Log_Shutdown(object sender, EventArgs e)
        {
        }

        #endregion

    }
}
