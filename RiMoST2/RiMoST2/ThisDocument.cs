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

namespace RiMoST2
{
    public partial class ThisDocument
    {
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            object noReset = false;
            object password = System.String.Empty;
            object useIRM = false;
            object enforceStyleLock = false;

            //TODO caricare dinamicamente i nomi che vistano il foglio


            /*NameValueCollection appSet = ConfigurationManager.AppSettings;

            string[] users = appSet.GetValues("utentiVisto")[0].Split(',');
            int rowNum = (int)Math.Ceiling(users.Length / 5.0);
            int colNum = (int)Math.Ceiling((decimal)users.Length / rowNum);

            Word.Table tb = this.Tables[1];

            tb.Rows[tb.Rows.Count].Cells.Split(rowNum, colNum);
            
            foreach (string usr in users)
            {
                for(int i = 0; i <

            }*/
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
