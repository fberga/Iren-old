using Iren.ToolsExcel.Utility;
using Iren.ToolsExcel.Base;
using System.Linq;
using System;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Iren.ToolsExcel.Forms
{
    public partial class FormSelezioneUP : Form
    {
        #region Variabili
        
        private string _siglaInformazione = "";
        private Dictionary<string, string> _upList = new Dictionary<string, string>();
        
        #endregion

        #region Proprietà

        public List<string> ListaUP
        {
            get 
            {
                return _upList.Keys.ToList();
            }
        }
        
        #endregion

        #region Costruttori

        public FormSelezioneUP(string siglaInformazione)
        {
            InitializeComponent();
            _siglaInformazione = siglaInformazione;
            this.Text = Simboli.nomeApplicazione + " - Selezione UP";

            DataView entitaInformazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
            entitaInformazioni.RowFilter = "SiglaInformazione = '" + _siglaInformazione + "'";

            string rowFilter = "SiglaEntita IN (";
            foreach (DataRowView entitaInfo in entitaInformazioni)
            {
                rowFilter += "'" + entitaInfo["SiglaEntita"] + "',";
            }
            rowFilter = rowFilter.Substring(0, rowFilter.Length - 1) + ")";

            DataView categorieEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            categorieEntita.RowFilter = rowFilter;

            _upList =
                (from r in categorieEntita.ToTable(true, "SiglaEntita", "DesEntita").AsEnumerable()
                 select r).ToDictionary(r => r["SiglaEntita"].ToString(), r => r["DesEntita"].ToString());

            comboUP.DataSource = new BindingSource(_upList, null);
            comboUP.DisplayMember = "Value";
            comboUP.ValueMember = "Key";
            comboUP.SelectedIndex = 0;
        }

        #endregion

        #region Eventi

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            comboUP.SelectedIndex = -1;
            this.Close();
        }

        private void btnCarica_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Sposta la selezione sul titolo dell'UP scelta e ritorna la sua sigla.
        /// </summary>
        /// <returns>Restituisce la sigla dell'UP scelta.</returns>
        public new object ShowDialog()
        {
            base.ShowDialog();

            if (comboUP.SelectedIndex != -1) 
            {
                //non mi serve il nome del foglio perché lavoro direttamente con la siglaEntita
                NewDefinedNames n = new NewDefinedNames("", NewDefinedNames.InitType.GOTOsOnly);
                string address = n.GetGotoFromSiglaEntita(comboUP.SelectedValue);
                Handler.Goto(address);
            }

            return comboUP.SelectedValue;
        } 

        #endregion
    }
}
