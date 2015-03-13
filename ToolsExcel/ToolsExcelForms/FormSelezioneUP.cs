using Iren.ToolsExcel.Utility;
using Iren.ToolsExcel.Base;
using System;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Forms
{
    public partial class FormSelezioneUP : Form
    {
        #region Variabili
        
        private string _siglaEntita = "";
        private bool _isCanceld = false;
        private bool _hasSelection = false;
        private string _siglaInformazione = "";
        
        #endregion

        #region Proprietà

        public bool IsCanceld { get { return _isCanceld; } }
        public bool HasSelection { get { return _hasSelection; } }
        public string SiglaEntita { get { return _siglaEntita; } }
        
        #endregion

        #region Costruttori

        public FormSelezioneUP(string siglaInformazione)
        {
            InitializeComponent();

            _siglaInformazione = siglaInformazione;

            this.Text = Simboli.nomeApplicazione + " - Selezione UP";
        }

        #endregion

        #region Eventi

        private void frmSELUP_Load(object sender, EventArgs e)
        {
            DataView entitaInformazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            entitaInformazioni.RowFilter = "SiglaInformazione = '" + _siglaInformazione + "'";

            string rowFilter = "SiglaEntita IN (";
            foreach (DataRowView entitaInfo in entitaInformazioni)
            {
                rowFilter += "'" + entitaInfo["SiglaEntita"] + "',";
            }
            rowFilter = rowFilter.Substring(0, rowFilter.Length - 1) + ")";

            DataView categorieEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            categorieEntita.RowFilter = rowFilter;

            DataView groupedEntita = categorieEntita.ToTable(true, "SiglaEntita", "DesEntita").DefaultView;

            comboUP.DataSource = groupedEntita;
            comboUP.DisplayMember = "DesEntita";
            comboUP.SelectedIndex = 0;
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            _siglaEntita = "";
            this.Close();
        }

        private void btnCarica_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void comboUP_SelectedIndexChanged(object sender, EventArgs e)
        {
            _siglaEntita = ((DataRowView)comboUP.SelectedItem)["SiglaEntita"].ToString();
        }

        /// <summary>
        /// Sposta la selezione sul titolo dell'UP scelta e ritorna la sua sugla.
        /// </summary>
        /// <returns>Restituisce la sigla dell'UP scelta.</returns>
        public new string ShowDialog()
        {
            base.ShowDialog();

            if(_siglaEntita != "")
                Handler.GOTO(_siglaEntita);            

            return _siglaEntita;
        } 

        #endregion
    }
}
