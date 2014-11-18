﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Iren.FrontOffice.Core;

namespace RiMoST2
{
    public partial class FormAnnullaModifica : Form
    {
        public FormAnnullaModifica()
        {
            InitializeComponent();
        }

        private void FormAnnullaModifica_Load(object sender, EventArgs e)
        {
            QryParams parameters = new QryParams() 
            {
                {"@IdRichiesta", "all"}
            };

            DataView dv = ThisDocument._db.Select("spGetRichiesta", parameters).DefaultView;
            dv.RowFilter = "IdTipologiaStato NOT IN (4, 7) AND IdRichiesta LIKE '%" + Globals.Ribbons.RiMoST.cbAnniDisponibili.Text + "'";
            cmbRichiesta.DataSource = dv;
            cmbRichiesta.DisplayMember = "IdRichiesta";
            //cmbRichiesta_SelectedIndexChanged(null, new EventArgs());
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Sei sicuro di voler ANNULLARE la richiesta selezionata?", "Attenzione!", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.OK)
            {
                DataRowView row = (DataRowView)cmbRichiesta.SelectedItem;

                QryParams parameters = new QryParams() 
                {
                    {"@IdRichiesta", row["IdRichiesta"]}
                };
                try
                {
                    ThisDocument._db.Insert("spAnnullaRichiesta", parameters);
                }
                catch (Exception)
                {
                    MessageBox.Show("Errore nell'annullamento della richiesta. Riporvare più tardi.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            this.Close();
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmbRichiesta_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataRowView row = (DataRowView)cmbRichiesta.SelectedItem;
            string path = @"file:///" + row["NomeFile"];
            DocPreview.Navigate(path);
        }
    }
}
