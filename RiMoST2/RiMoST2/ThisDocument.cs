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
using Iren.FrontOffice.Core;

namespace RiMoST2
{
    public partial class ThisDocument
    {
        #region Variabili

        public static DataBase _db;

        #endregion

        #region Callbacks

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            Globals.ThisDocument.pictureBox1.Visible = false;
            Connection.CryptSection(System.Reflection.Assembly.GetExecutingAssembly());

            NameValueCollection appSet = ConfigurationManager.AppSettings;

            string[] users = appSet.GetValues("utentiVisto")[0].Split(',');
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

            DataView dtView = _db.Select("spGetApplicazioniDisponibili").DefaultView;

            cmbStrumento.DataSource = dtView;
            cmbStrumento.DisplayMember = "DesApplicazione";            

            DataTable dt = _db.Select("spGetFirstAvailableID");
            lbIdRichiesta.Text = dt.Rows[0][0].ToString();

            lbDataInvio.Text = DateTime.Now.ToShortDateString();

            txtDescrizione.Multiline = true;
            txtDescrizione.Height = 265.5f;
            txtNote.Multiline = true;
            txtNote.Height = 54.75f;
            txtOggetto.Multiline = true;
            txtOggetto.Height = 33f;

            object what = Word.WdGoToItem.wdGoToLine;
            object which = Word.WdGoToDirection.wdGoToLast;
            object missing = Missing.Value;

            AddProtection();
        }
        private void ThisDocument_BeforeClose(object sender, System.ComponentModel.CancelEventArgs e)
        {
            CloseWithoutSaving();
        }
        #pragma warning disable 0467
        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
            Application.Quit();
        }
        #pragma warning restore 0467

        #endregion

        #region Metodi

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
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);
            this.BeforeClose += new System.ComponentModel.CancelEventHandler(this.ThisDocument_BeforeClose);

        }

        #endregion


    }
}
