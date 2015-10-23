﻿using System;
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
using System.Configuration;
using System.IO;
using System.Deployment.Application;
using System.Reflection;
using System.Globalization;
using Iren.ToolsExcel.Utility;
using Iren.ToolsExcel.Base;

// ************************************************************* SISTEMA COMANDI ************************************************************* //

namespace Iren.ToolsExcel
{
    public partial class ThisWorkbook : IToolsExcelThisWorkbook
    {
        #region Proprietà

        public System.Version Version 
        { 
            get 
            {
                try
                {
                    return ApplicationDeployment.CurrentDeployment.CurrentVersion;
                }
                catch (Exception)
                {
                    return Assembly.GetExecutingAssembly().GetName().Version;
                }
            }
        }

        public Worksheet Main { get { return Globals.Main.Base; } }
        public Worksheet Log { get { return Globals.Log.Base; } }
        public new Worksheet ActiveSheet { get { return (Worksheet)base.ActiveSheet; } }

        public int IdApplicazione { get { return idApplicazione; } set { idApplicazione = value; } }
        public int IdUtente { get { return idUtente; } set { idUtente = value; } }
        public string NomeUtente { get { return nomeUtente; } set { nomeUtente = value; } }
        public DateTime DataAttiva { get { return dataAttiva; } set { dataAttiva = value; } }
        public string Ambiente { get { return ambiente; } set { ambiente = value; } }
        public string Password { get { return password; } }

        public DataSet RepositoryDataSet { get { return repositoryDataSet; } }
        public DataSet LogDataSet { get { return logDataSet; } }
        public DataSet RibbonDataSet { get { return ribbonDataSet; } }

        #endregion

        #region Cached Attribute

        [CachedAttribute()]
        public int idApplicazione = 8;
        [CachedAttribute()]
        public int idUtente = -1;
        [CachedAttribute()]
        public string nomeUtente = string.Empty;
        [CachedAttribute()]
        public DateTime dataAttiva = DateTime.Now;
        [CachedAttribute()]
        public string ambiente = Simboli.PROD;
        [CachedAttribute()]
        public DataSet repositoryDataSet = new DataSet();
        [CachedAttribute()]
        public DataSet logDataSet = new DataSet();
        [CachedAttribute()]
        public DataSet ribbonDataSet = new DataSet();
        [CachedAttribute()]
        public string password = "8176";

        #endregion

        #region Codice generato dalla finestra di progettazione di VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.BeforeClose += new Microsoft.Office.Interop.Excel.WorkbookEvents_BeforeCloseEventHandler(this.ThisWorkbook_BeforeClose);
            this.SheetSelectionChange += new Microsoft.Office.Interop.Excel.WorkbookEvents_SheetSelectionChangeEventHandler(Handler.CellClick);
            this.Startup += new System.EventHandler(this.ThisWorkbook_Startup);
        }

        #endregion

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
#if DEBUG
            ambiente = Simboli.DEV;
#endif
            //Utility.Workbook.StartUp(this);      
        }
        private void ThisWorkbook_BeforeClose(ref bool Cancel)
        {
            Utility.Workbook.Close();
            Save();
        }
    }
}
