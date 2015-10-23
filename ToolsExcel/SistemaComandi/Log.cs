﻿using System;
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
using Iren.ToolsExcel.Utility;
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

        }

        #endregion

        #region Callbacks

        private void Log_Startup(object sender, EventArgs e)
        {
            //Globals.ThisWorkbook.LogDataSet.Tables[DataBase.Tab.LOG];



            //ListObject logObj = Controls.AddListObject(this.Range["A1"], "LogList");
            //logObj.TableStyle = "TableStyleLight16";

            //logObj.AutoSetDataBoundColumnHeaders = true;
            //logObj.DataSource = Globals.ThisWorkbook.logDataSet;





            //logObj.DataMember = DataBase.Tab.LOG;

            //if (!Simboli.Aborted)
            //{
            //    Unprotect(Simboli.pwd);
            //    try
            //    {
            //        _logObj = Globals.Factory.GetVstoObject(ListObjects["LogList"]);
            //    }
            //    catch (Exception)
            //    {
            //        _logObj = Controls.AddListObject(Range["A1"], "LogList");
            //    }
            //    _logObj.AutoSetDataBoundColumnHeaders = true;

            //    if (DataBase.OpenConnection())
            //    {
            //        _logObj.DataSource = DataBase.LocalDB.Tables[DataBase.Tab.LOG].DefaultView;
            //        _logObj.Range.EntireColumn.AutoFit();
            //        _logObj.TableStyle = "TableStyleLight16";

            //        Excel.Range rng = Columns[2];
            //        rng.NumberFormat = "dd/MM/yyyy";
            //        rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            //        DataBase.CloseConnection();
            //    }
            //    Protect(Simboli.pwd);
            //}
        }

        #endregion

    }
}
