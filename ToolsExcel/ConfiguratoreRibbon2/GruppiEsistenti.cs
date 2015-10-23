using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public partial class GruppiEsistenti : Form
    {
        DataTable _allGroups;
        DataTable _allFunctions;
        Control _ribbon;

        public GruppiEsistenti(Control ribbon)
        {
            InitializeComponent();
            _ribbon = ribbon;

            _allGroups = DataBase.Select(SP.GRUPPO_CONTROLLO, "@IdApplicazione=-1;@IdUtente=-1");

            var unusedGroups = _allGroups.AsEnumerable()
                .Where(r => !ConfiguratoreRibbon.GruppoControlloUtilizzati.Contains((int)r["IdGruppoControllo"]))
                .Where(r => !ConfiguratoreRibbon.GruppiUtilizzati.Contains((int)r["IdGruppo"]))
                .Select(r => new { LabelGruppo = r["LabelGruppo"], IdGruppo = r["IdGruppo"] })
                .Distinct()                
                .ToList();

            listBoxGruppi.DisplayMember = "LabelGruppo";
            listBoxGruppi.ValueMember = "IdGruppo";

            _allFunctions = DataBase.Select(SP.CONTROLLO_FUNZIONE);
            DataView funzioni = _allFunctions.DefaultView;

            funzioni.RowFilter = "IdFunzione=-1";
            listBoxFunzioni.DisplayMember = "NomeFunzione";
            listBoxFunzioni.ValueMember = "IdFunzione";

            listBoxFunzioni.DataSource = funzioni;
            listBoxGruppi.DataSource = unusedGroups;
        }

        private void CambioGruppo(object sender, EventArgs e)
        {
            if (listBoxGruppi.SelectedValue != null)
            {
                

                var users = _allGroups.AsEnumerable()
                    .Where(r => r["IdGruppo"].Equals(listBoxGruppi.SelectedValue))
                    .Select(r => new { IdUtente = r["IdUtente"]})
                    .OrderBy(r => r.IdUtente)
                    .Distinct()
                    .ToList();

                

                listBoxUtenti.DataSource = users;
                listBoxUtenti.ValueMember = "IdUtente";
                listBoxUtenti.DisplayMember = "IdUtente";
            }
        }

        private void CambioApplicazione(object sender, EventArgs e)
        {
            if (listBoxApplicazioni.SelectedValue != null)
            {
                //carico anteprima gruppo
                var controls = _allGroups.AsEnumerable()
                    .Where(r => r["IdGruppo"].Equals(listBoxGruppi.SelectedValue) && r["IdApplicazione"].Equals(listBoxApplicazioni.SelectedValue) && r["IdUtente"].Equals(listBoxUtenti.SelectedValue))
                    .ToList();

                RibbonGroup grp = new RibbonGroup(panelRibbonLayout, (int)listBoxGruppi.SelectedValue);
                grp.Text = listBoxGruppi.Text;
                panelRibbonLayout.Controls.Clear();
                panelRibbonLayout.Controls.Add(grp);

                foreach (DataRow r in controls)
                {
                    Control ctrl = Utility.AddControlToGroup(grp, r, _allFunctions);
                    ctrl.GotFocus += EvidenziaFunzioni;
                    ctrl.Tag = r["IdGruppoControllo"];
                }
            }
        }

        private void CambioUtente(object sender, EventArgs e)
        {
            if (listBoxUtenti.SelectedValue != null)
            {
                var applications = _allGroups.AsEnumerable()
                    .Where(r => r["IdGruppo"].Equals(listBoxGruppi.SelectedValue) && r["IdUtente"].Equals(listBoxUtenti.SelectedValue))
                    .Select(r => new { IdApplicazione = r["IdApplicazione"], DesApplicazione = r["DesApplicazione"] })
                    .Distinct()
                    .ToList();

                listBoxApplicazioni.DataSource = applications;
                listBoxApplicazioni.ValueMember = "IdApplicazione";
                listBoxApplicazioni.DisplayMember = "DesApplicazione";
            }
        }

        private void EvidenziaFunzioni(object sender, EventArgs e)
        {
            IRibbonControl ctrl = sender as IRibbonControl;
            if (ctrl.Functions.Count > 0)
                _allFunctions.DefaultView.RowFilter = "IdGruppoControllo = " + ((Control)ctrl).Tag + " AND IdFunzione IN (" + string.Join(",", ctrl.Functions) + ")";
            else
                _allFunctions.DefaultView.RowFilter = "IdFunzione = -1";
        }

        private void AggiungiGruppo_Click(object sender, EventArgs e)
        {
            var ribbonGroup = panelRibbonLayout.Controls.OfType<RibbonGroup>().First();
            var ctrls = Utility.GetAll(ribbonGroup);

            foreach (Control ctrl in ctrls)
                ctrl.GotFocus -= EvidenziaFunzioni;

            Utility.AddGroupToRibbon(_ribbon, ribbonGroup);
            
            //listBoxGruppi.Items.RemoveAt(listBoxGruppi.SelectedIndex);
        }
    }
}
