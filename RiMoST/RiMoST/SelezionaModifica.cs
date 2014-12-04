using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Iren.FrontOffice.Core;
using System.Globalization;
using System.Configuration;

namespace Iren.FrontOffice.Tools
{
    public partial class SelezionaModifica : Form
    {
        #region Variabili
        string _anno;
        public bool _chkIsDraft;
        public bool _btnRefreshEnabled;
        #endregion

        #region Costruttori
        public SelezionaModifica(string anno, bool chkIsDraft, bool btnRefreshEnabled)
        {
            _anno = anno;
            _chkIsDraft = chkIsDraft;
            _btnRefreshEnabled = btnRefreshEnabled;
            InitializeComponent();
        }
        #endregion

        #region Callbacks
        private void SelezionaModifica_Load(object sender, EventArgs e)
        {

            DataView dv = ThisDocument._db.Select("spGetRichiesta", "@IdTipologiaStato:7; @IdStruttura:" + ThisDocument._idStruttura).DefaultView;
            dv.RowFilter = "IdRichiesta LIKE '%" + _anno + "'";
            cmbRichiesta.DataSource = dv;
            cmbRichiesta.DisplayMember = "IdRichiesta";
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (cmbRichiesta.Text == "")
            {
                MessageBox.Show("Non ci sono bozze al momento.", "Attenzione!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                DataRowView row = (DataRowView)cmbRichiesta.SelectedItem;
                Globals.ThisDocument.lbIdRichiesta.Text = "" + row["IdRichiesta"];
                ((DataView)Globals.ThisDocument.cmbStrumento.DataSource).RowFilter = "IdApplicazione = " + row["IdApplicazione"];
                Globals.ThisDocument.cmbStrumento.Enabled = false;
                Globals.ThisDocument.txtOggetto.Text = "" + row["Oggetto"];
                Globals.ThisDocument.txtDescrizione.Text = "" + row["Descr"];
                Globals.ThisDocument.txtNote.Text = "" + row["Note"];
                _chkIsDraft = true;
                _btnRefreshEnabled = false;
                this.Hide();
            }
        }
        #endregion
    }
}
