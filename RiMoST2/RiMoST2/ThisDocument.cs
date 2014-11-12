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

namespace RiMoST2
{
    public partial class ThisDocument
    {
        static DataBase _db;

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            object noReset = false;
            object password = System.String.Empty;
            object useIRM = false;
            object enforceStyleLock = false;

            //TODO caricare dinamicamente i nomi che vistano il foglio

            NameValueCollection appSet = ConfigurationManager.AppSettings;

            string[] users = appSet.GetValues("utentiVisto")[0].Split(',');
            int rowNum = (int)Math.Ceiling(users.Length / 5.0);
            int colNum = (int)Math.Ceiling((decimal)users.Length / rowNum);

            Word.Table tb = this.Tables[1];

            tb.Rows[tb.Rows.Count].Cells.Split(rowNum, colNum);
            

            int i = 0;
            int j = 0;
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
            txtDescrizione.Height = 199.5f;
            txtNote.Multiline = true;
            txtNote.Height = 99.75f;
            txtOggetto.Multiline = true;
            txtOggetto.Height = 34.5f;

            this.Protect(Word.WdProtectionType.wdAllowOnlyFormFields,
                ref noReset, ref password, ref useIRM, ref enforceStyleLock);
        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Codice generato dalla finestra di progettazione di VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion
    }
}
