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

                btnRimuoviParD.Enabled = 
                    (from r1 in valCorretti.AsEnumerable()
                     where (DateTime)r1["Inizio Validità"] > DateTime.Today
                     select r1).Count() > 0;

                dataGridParametriD.DataSource = valCorretti;
                dataGridParametriD.Columns["Fine Validità"].DefaultCellStyle.NullValue = "-";
                dataGridParametriD.Columns["Valore"].DefaultCellStyle.Format = "0.#########";
            }
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

                btnRimuoviParH.Enabled =
                    (from r1 in valCorretti.AsEnumerable()
                     where (DateTime)r1["Inizio Validità"] > DateTime.Today
                     select r1).Count() > 0;

                dataGridParametriH.DataSource = valCorretti;
                dataGridParametriH.Columns["Fine Validità"].DefaultCellStyle.NullValue = "-";
                dataGridParametriH.Columns["Valore"].DefaultCellStyle.Format = "0.#########";
            }
        }
            
    }
}
