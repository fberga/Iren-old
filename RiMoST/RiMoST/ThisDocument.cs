using System;
using System.Configuration;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Specialized;
using DataTable = System.Data.DataTable;
using DataRow = System.Data.DataRow;
using DataView = System.Data.DataView;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using Iren.ToolsExcel.Core;
using System.IO;

namespace Iren.FrontOffice.Tools
{
    public partial class ThisDocument
    {
        #region Variabili

        public static DataBase _db;
        public static string _idStruttura;

        #endregion

        #region Callbacks

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            SetAppConfigEnvironment();
            _idStruttura = ConfigurationManager.AppSettings["idStruttura"];
            string[] users = ConfigurationManager.AppSettings["utentiVisto"].Split(',');
            int rowNum = (int)Math.Ceiling(users.Length / 5.0);
            int colNum = (int)Math.Ceiling((decimal)users.Length / rowNum);

            Word.Table tb = this.Tables[1];

            tb.Rows[tb.Rows.Count].Cells.Split(rowNum, colNum);

            int i = 0, j = 0;            
            foreach (string usr in users)
            {
                tb.Cell((tb.Rows.Count - rowNum) + i + 1, (j % colNum) + 1).Range.Text = usr;

                if ((++j % colNum) == 0)
                    i++;
            }

            _db = new DataBase(ConfigurationManager.AppSettings["DB"]);
            if(_db.OpenConnection())
            {
                DataView dtView = _db.Select("spGetApplicazioniDisponibili", "@IdStruttura=" + _idStruttura).DefaultView;

                cmbStrumento.DataSource = dtView;
                cmbStrumento.DisplayMember = "DesApplicazione";

                DataTable dt = _db.Select("spGetFirstAvailableID", "@IdStruttura=" + _idStruttura);
                lbIdRichiesta.Text = dt.Rows[0][0].ToString();

                lbDataInvio.Text = DateTime.Now.ToShortDateString();

                object what = Word.WdGoToItem.wdGoToLine;
                object which = Word.WdGoToDirection.wdGoToLast;
                object missing = Missing.Value;

                _db.CloseConnection();
            }

            AddProtection();
        }
        private void ThisDocument_BeforeClose(object sender, System.ComponentModel.CancelEventArgs e)
        {
            CloseWithoutSaving();
        }
        #pragma warning disable 0467
        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
            ThisApplication.Quit();
        }
        #pragma warning restore 0467

        #endregion

        #region Metodi

        private void SetAppConfigEnvironment()
        {
            string file = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RiMoST/RiMoST.config");
            if (File.Exists(AppDomain.CurrentDomain.GetData("APP_CONFIG_FILE").ToString()))
            {
                Directory.CreateDirectory(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RiMoST"));
                File.Copy(AppDomain.CurrentDomain.GetData("APP_CONFIG_FILE").ToString(), file, true);
                File.Delete(AppDomain.CurrentDomain.GetData("APP_CONFIG_FILE").ToString());
            }
            AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", file);
            CryptHelper.CryptSection("connectionStrings");
        }

        public void AddProtection()
        {
            object noReset = false;
            object password = System.String.Empty;
            object useIRM = false;
            object enforceStyleLock = false;

            this.Protect(Word.WdProtectionType.wdAllowOnlyFormFields,
                ref noReset, ref password, ref useIRM, ref enforceStyleLock);
        }
        public void RemoveProtection()
        {
            object password = System.String.Empty;
            this.Unprotect(ref password);
        }

        public void CloseWithoutSaving()
        {
            object saveMod = WdSaveOptions.wdDoNotSaveChanges;
            object missing = Missing.Value;
            this.Close(ref saveMod, ref missing, ref missing);
        }

        public static void Highlight(string textToFind, Word.WdColorIndex color, string highlightMark = "")
        {
            Word.Find finder = Globals.ThisDocument.Content.Find;
            finder.Text = textToFind;
            finder.Replacement.Text = finder.Text + highlightMark;
            finder.Replacement.Font.ColorIndex = color;
            finder.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }
        public static void ToNormal(string textToFind, Word.WdColorIndex color, string highlightMark = "")
        {
            Word.Find finder = Globals.ThisDocument.Content.Find;
            finder.Text = textToFind + highlightMark;
            finder.Replacement.Text = textToFind;
            finder.Replacement.Font.ColorIndex = color;
            finder.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }

        #endregion

        #region Codice generato dalla finestra di progettazione di VSTO

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.txtDescrizione.Entering += new Microsoft.Office.Tools.Word.ContentControlEnteringEventHandler(this.TextArea_Entering);
            this.txtDescrizione.Exiting += new Microsoft.Office.Tools.Word.ContentControlExitingEventHandler(this.TextArea_Exiting);
            this.txtOggetto.Entering += new Microsoft.Office.Tools.Word.ContentControlEnteringEventHandler(this.TextArea_Entering);
            this.txtOggetto.Exiting += new Microsoft.Office.Tools.Word.ContentControlExitingEventHandler(this.TextArea_Exiting);
            this.txtNote.Entering += new Microsoft.Office.Tools.Word.ContentControlEnteringEventHandler(this.TextArea_Entering);
            this.txtNote.Exiting += new Microsoft.Office.Tools.Word.ContentControlExitingEventHandler(this.TextArea_Exiting);
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);
            this.BeforeClose += new System.ComponentModel.CancelEventHandler(this.ThisDocument_BeforeClose);

        }

        #endregion

        private void TextArea_Entering(object sender, ContentControlEnteringEventArgs e)
        {
            RemoveProtection();
        }

        private void TextArea_Exiting(object sender, ContentControlExitingEventArgs e)
        {
            AddProtection();
        }


    }
}
