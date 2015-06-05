using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Iren.ToolsExcel.Base;
using System.Globalization;


namespace Iren.ToolsExcel.Forms
{
    public partial class FormModificaParametri : Form
    {
        DataView _parametriD = new DataView();
        DataView _parametriH = new DataView();
        DataTable _parametri;
        DataView _entita;

        public FormModificaParametri()
        {
            InitializeComponent();

            this.Text = Simboli.nomeApplicazione + " - Modifica Parametri";

            _entita = new DataView(Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.CATEGORIA_ENTITA]);
            _parametri = Utility.DataBase.Select(Utility.DataBase.SP.ELENCO_PARAMETRI, "@IdApplicazione=" + Utility.DataBase.DB.IdApplicazione);

            _parametriD = new DataView(_parametri);
            _parametriH = new DataView(_parametri);

            cmbEntita.ValueMember = "SiglaEntita";
            cmbEntita.DisplayMember = "DesEntita";

            cmbParametriD.DisplayMember = "Descrizione";

            cmbParametriH.DisplayMember = "Descrizione";

            cmbEntita.DataSource = _entita;
            cmbParametriD.DataSource = _parametriD;
            cmbParametriH.DataSource = _parametriH;
        }

        private void cmbEntita_SelectedIndexChanged(object sender, EventArgs e)
        {
            _parametriD.RowFilter = "SiglaEntita = '" + cmbEntita.SelectedValue + "' AND Dettaglio = 'D'";
            _parametriH.RowFilter = "SiglaEntita = '" + cmbEntita.SelectedValue + "' AND Dettaglio = 'H'";

            if (_parametriD.Count == 0)
                ((Control)tabPageParD).Enabled = false;
            else
                ((Control)tabPageParD).Enabled = true;

            if (_parametriH.Count == 0)
                ((Control)tabPageParH).Enabled = false;
            else
                ((Control)tabPageParH).Enabled = true;

            cmbParametriH_SelectedIndexChanged(null, null);
            cmbParametriD_SelectedIndexChanged(null, null);
        }

        private void cmbParametriH_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbParametriH.SelectedValue != null)
            {
                DataRowView r = cmbParametriH.SelectedValue as DataRowView;

                DataTable valori = Utility.DataBase.Select(Utility.DataBase.SP.VALORI_PARAMETRI, new Core.QryParams() 
                {
                    {"@IdApplicazione", Utility.DataBase.DB.IdApplicazione},
                    {"@IdEntita", r["IdEntita"]},
                    {"@IdTipologiaParametro", r["IdParametro"]},
                    {"@Dettaglio", "H"},
                });

                DataTable valCorretti = new DataTable()
                {
                    Columns = 
                    {
                        {"Inizio Validità", typeof(DateTime)},
                        {"Fine Validità", typeof(DateTime)},
                        {"Ora", typeof(int)},
                        {"Valore", typeof(decimal)}
                    }
                };


                foreach (DataRow val in valori.Rows)
                {
                    DataRow newRow = valCorretti.NewRow();

                    DateTime fineValidita = DateTime.ParseExact(val["DataFV"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);

                    newRow["Inizio Validità"] = DateTime.ParseExact(val["DataIV"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                    if (fineValidita.Year < 6000)
                        newRow["Fine Validità"] = fineValidita;
                    newRow["Ora"] = int.Parse(val["Ora"].ToString());
                    newRow["Valore"] = decimal.Parse(val["Valore"].ToString(), CultureInfo.CurrentUICulture);

                    valCorretti.Rows.Add(newRow);
                }

                dataGridParametriH.DataSource = valCorretti;
                dataGridParametriH.Columns["Fine Validità"].DefaultCellStyle.NullValue = "-";
                dataGridParametriH.Columns["Valore"].DefaultCellStyle.Format = "0.#########";
            }
            else
            {
                dataGridParametriH.DataSource = null;
            }
        }
        private void cmbParametriD_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbParametriD.SelectedValue != null)
            {
                DataRowView r = cmbParametriD.SelectedValue as DataRowView;

                DataTable valori = Utility.DataBase.Select(Utility.DataBase.SP.VALORI_PARAMETRI, new Core.QryParams() 
                {
                    {"@IdApplicazione", Utility.DataBase.DB.IdApplicazione},
                    {"@IdEntita", r["IdEntita"]},
                    {"@IdTipologiaParametro", r["IdParametro"]},
                    {"@Dettaglio", "D"},
                });

                DataTable valCorretti = new DataTable()
                {
                    Columns = 
                {
                    {"Inizio Validità", typeof(DateTime)},
                    {"Fine Validità", typeof(DateTime)},
                    {"Valore", typeof(decimal)}
                }
                };


                foreach (DataRow val in valori.Rows)
                {
                    DataRow newRow = valCorretti.NewRow();

                    DateTime fineValidita = DateTime.ParseExact(val["DataFV"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);

                    newRow["Inizio Validità"] = DateTime.ParseExact(val["DataIV"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                    if (fineValidita.Year < 6000)
                        newRow["Fine Validità"] = fineValidita;
                    newRow["Valore"] = decimal.Parse(val["Valore"].ToString(), CultureInfo.CurrentUICulture);

                    valCorretti.Rows.Add(newRow);
                }

                dataGridParametriD.DataSource = valCorretti;
                dataGridParametriD.Columns["Fine Validità"].DefaultCellStyle.NullValue = "-";
                dataGridParametriD.Columns["Valore"].DefaultCellStyle.Format = "0.#########";
            }
            else
            {
                dataGridParametriD.DataSource = null;
            }
        }

        private void DataGridMouseDown(DataGridView sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                var info = sender.HitTest(e.X, e.Y);

                if (info.RowIndex >= 0)
                {
                    if ((DateTime)sender["Inizio Validità", info.RowIndex].Value > DateTime.Today)
                    {
                        modificareValoreToolStripMenuItem.Enabled = true;
                        cancellaParametroToolStripMenuItem.Enabled = true;
                    }
                    else
                    {
                        modificareValoreToolStripMenuItem.Enabled = false;
                        cancellaParametroToolStripMenuItem.Enabled = false;
                    }
                }
                else
                {
                    modificareValoreToolStripMenuItem.Enabled = false;
                    cancellaParametroToolStripMenuItem.Enabled = false;
                }
            }
        }
        private void dataGridParametriH_MouseDown(object sender, MouseEventArgs e)
        {
            DataGridMouseDown(dataGridParametriH, e);
        }
        private void dataGridParametriD_MouseDown(object sender, MouseEventArgs e)
        {
            DataGridMouseDown(dataGridParametriD, e);
        }

        private void DataGridMouseMove(DataGridView sender, MouseEventArgs e)
        {
            var info = sender.HitTest(e.X, e.Y);

            if (info.RowIndex >= 0 && !_onEdit)
            {
                if (!sender.Rows[info.RowIndex].Selected)
                    sender.ClearSelection();

                sender.CurrentCell = sender.Rows[info.RowIndex].Cells[0];
                sender.CurrentRow.Selected = true;
            }
        }
        private void dataGridParametriH_MouseMove(object sender, MouseEventArgs e)
        {
            DataGridMouseMove(dataGridParametriH, e);
        }
        private void dataGridParametriD_MouseMove(object sender, MouseEventArgs e)
        {
            DataGridMouseMove(dataGridParametriD, e);
        }

        bool _fromMenuStrip = false;
        private void modificareValoreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var source = (DataGridView)contextMenuDataGrid.SourceControl;

            int index = source.SelectedRows[0].Index;
            source.CurrentCell = source.SelectedRows[0].Cells["Valore"];
            _fromMenuStrip = true;
            source.BeginEdit(false);
        }

        bool _onEdit = false;
        private void dataGridCellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (_fromMenuStrip)
            {
                _onEdit = true;
                _fromMenuStrip = false;
            }
            else
                e.Cancel = true;
        }
        private void dataGridCellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            _onEdit = false;
            cmbParametriH_SelectedIndexChanged(null, null);
            cmbParametriD_SelectedIndexChanged(null, null);
        }

        private void dataGridParametriH_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (_onEdit)
            {
                decimal newValue;
                if (decimal.TryParse(e.FormattedValue.ToString(), NumberStyles.AllowDecimalPoint, CultureInfo.InstalledUICulture, out newValue))
                {
                    DataRowView parRow = (DataRowView)cmbParametriH.SelectedValue;
                    DataRow valueRow = ((DataTable)dataGridParametriH.DataSource).Rows[e.RowIndex];

                    if (!Utility.DataBase.Insert(Utility.DataBase.SP.UPDATE_PARAMETRO, new Core.QryParams()
                        {
                            {"@IdEntita", parRow["IdEntita"]},
                            {"@IdTipologiaParametro", parRow["IdParametro"]},
                            {"@DataIV", ((DateTime)valueRow["Inizio Validità"]).ToString("yyyyMMdd")},
                            {"@DataFV", ((DateTime)valueRow["Fine Validità"]).ToString("yyyyMMdd")},
                            {"@Ora", ((int)valueRow["Ora"]).ToString("00")},
                            {"@Valore", newValue},
                            {"@Dettaglio", "H"}
                        })) 
                    {
                        MessageBox.Show("Ci sono stati problemi nel salvataggio della modifica... Riprovare più tardi.", Simboli.nomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        e.Cancel = true;
                    };
                }
                else
                {
                    MessageBox.Show("Il nuovo valore inserito presenta dei caratteri non validi.", Simboli.nomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    e.Cancel = true;
                }
            }
        }
        private void dataGridParametriD_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (_onEdit)
            {
                decimal newValue;
                if (decimal.TryParse(e.FormattedValue.ToString(), NumberStyles.AllowDecimalPoint, CultureInfo.InstalledUICulture, out newValue))
                {
                    DataRowView parRow = (DataRowView)cmbParametriH.SelectedValue;
                    DataRow valueRow = ((DataTable)dataGridParametriH.DataSource).Rows[e.RowIndex];

                    if (!Utility.DataBase.Insert(Utility.DataBase.SP.UPDATE_PARAMETRO, new Core.QryParams()
                        {
                            {"@IdEntita", parRow["IdEntita"]},
                            {"@IdTipologiaParametro", parRow["IdParametro"]},
                            {"@DataIV", ((DateTime)valueRow["Inizio Validità"]).ToString("yyyyMMdd")},
                            {"@DataFV", ((DateTime)valueRow["Fine Validità"]).ToString("yyyyMMdd")},
                            {"@Valore", newValue},
                            {"@Dettaglio", "D"}
                        }))
                    {
                        MessageBox.Show("Ci sono stati problemi nel salvataggio della modifica... Riprovare più tardi.", Simboli.nomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        e.Cancel = true;
                    };
                }
                else
                {
                    MessageBox.Show("Il nuovo valore inserito presenta dei caratteri non validi.", Simboli.nomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    e.Cancel = true;
                }
            }
        }

        private void DataGridMouseEnter(object sender, EventArgs e)
        {
            ((DataGridView)sender).Select();
        }

        
    }
}
