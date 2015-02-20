using Iren.FrontOffice.Base;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Iren.FrontOffice.Forms
{
    public partial class FormSelezioneUP : Form
    {
        public bool _isDeleted = false;
        public bool _hasSelection = false;
        public string _siglaEntita;

        public FormSelezioneUP()
        {
            InitializeComponent();

            this.Text = Simboli.nomeApplicazione + " - Selezione UP";
        }

        private void frmSELUP_Load(object sender, EventArgs e)
        {
            DataView entitaInformazioni = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAINFORMAZIONE].DefaultView;
            entitaInformazioni.RowFilter = "SiglaInformazione = 'OTTIMO'";

            DataView categorieEntita = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.CATEGORIAENTITA].DefaultView;
            categorieEntita.RowFilter = "SiglaEntita = '" + entitaInformazioni[0]["SiglaEntita"] + "'";

            DataView groupedEntita = categorieEntita.ToTable(true, "SiglaEntita", "DesEntita").DefaultView;

            comboUP.DataSource = groupedEntita;
            comboUP.DisplayMember = "DesEntita";
            comboUP.SelectedIndex = 0;
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            _isDeleted = true;
            this.Close();
        }

        private void btnCarica_Click(object sender, EventArgs e)
        {
            _hasSelection = true;
            this.Close();
        }

        private void comboUP_SelectedIndexChanged(object sender, EventArgs e)
        {
            _siglaEntita = ((DataRowView)comboUP.SelectedItem)["SiglaEntita"].ToString();
        }
    }
}
