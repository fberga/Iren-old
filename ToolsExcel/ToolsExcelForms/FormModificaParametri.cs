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
        bool _fromMenuStrip = false;
        bool _onEdit = false;
        bool _isUpdate = false;
        bool _currentCellModified = false;
        DateTime _currDataIV;

        public FormModificaParametri()
        {
            InitializeComponent();

            this.Text = Simboli.nomeApplicazione + " - Modifica Parametri";

            _entita = new DataView(Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.CATEGORIA_ENTITA]);
            _parametri = Utility.DataBase.Select(Utility.DataBase.SP.ELENCO_PARAMETRI, "@IdApplicazione=" + Utility.DataBase.DB.IdApplicazione) ?? new DataTable();

            _parametriD = new DataView(_parametri);
            _parametriH = new DataView(_parametri);

            cmbEntita.ValueMember = "SiglaEntita";
            cmbEntita.DisplayMember = "DesEntita";

            cmbParametriD.DisplayMember = "Descrizione";
            cmbParametriH.DisplayMember = "Descrizione";

            cmbEntita.DataSource = _entita;
            cmbParametriD.DataSource = _parametriD;
            cmbParametriH.DataSource = _parametriH;

            //TODO rimuovere quando saranno utilizzati i parametri orari
            ((Control)tabPageParH).Visible = false;
        }

        private void cmbEntita_SelectedIndexChanged(object sender, EventArgs e)
        {
            _parametriD.RowFilter = "SiglaEntita = '" + cmbEntita.SelectedValue + "' AND Dettaglio = 'D'";
            _parametriH.RowFilter = "SiglaEntita = '" + cmbEntita.SelectedValue + "' AND Dettaglio = 'H'";

            if (_parametriD.Count == 0)
                ((Control)tabPageParD).Enabled = false;
            else
                ((Control)tabPageParD).Enabled = true;

            //TODO abilitare quando saranno utilizzati i parametri orari
            //if (_parametriH.Count == 0)
            //    ((Control)tabPageParH).Enabled = false;
            //else
            //    ((Control)tabPageParH).Enabled = true;

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
                }) ?? new DataTable();

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

        private bool CheckIfInsertBeforeAllowed(DataTable dt, int r)
        {
            if (dt.Rows.Count == 0)
                return false;

            DateTime currentIV = (DateTime)dt.Rows[r]["Inizio Validità"];
            DateTime precedingIV = r == 0 ? DateTime.MinValue : (DateTime)dt.Rows[r - 1]["Inizio Validità"];
            DateTime precedingFV = r == 0 ? DateTime.MinValue : (DateTime)dt.Rows[r - 1]["Fine Validità"];

            return
                currentIV > DateTime.Today                                  //posso inserire un parametro con IV >= oggi
                && precedingFV > DateTime.Today.AddDays(-1)                 //posso arretrare di 1 giorno la fine validità della riga sopra
                && precedingIV != precedingFV;                              //ho spazio per ridimensionare la fine validità della riga sopra
        }
        private bool CheckIfInsertAfterAllowed(DataTable dt, int r)
        {
            if (dt.Rows.Count == 0)
                return false;

            if (r == dt.Rows.Count - 1)
                return true;

            DateTime currentIV = (DateTime)dt.Rows[r]["Inizio Validità"];
            DateTime currentFV = (DateTime)dt.Rows[r]["Fine Validità"];
            DateTime subsequentIV = (DateTime)dt.Rows[r + 1]["Inizio Validità"];

            return
                subsequentIV > DateTime.Today                               //posso inserire un parametro con IV >= oggi
                && currentFV > DateTime.Today.AddDays(-1)                   //posso arretrare di 1 giorno la fine validità della riga corrente
                && currentIV < currentFV;                                   //ho spazio per ridimensionare la fine validità della riga corrente
        }
        private void RefreshMenuItems()
        {
            if (dataGridParametriD.CurrentRow != null || dataGridParametriD.IsCurrentRowDirty)
            {                
                int index = dataGridParametriD.CurrentRow.Index;
                if (index >= 0)
                {
                    if (CheckIfInsertAfterAllowed((DataTable)dataGridParametriD.DataSource, index))
                    {
                        inserisciSottoContextMenu.Enabled = true;
                        inserisciSottoTopMenu.Enabled = true;
                    }
                    else
                    {
                        inserisciSottoContextMenu.Enabled = false;
                        inserisciSottoTopMenu.Enabled = false;
                    }

                    if (CheckIfInsertBeforeAllowed((DataTable)dataGridParametriD.DataSource, index))
                    {
                        inserisciSopraContextMenu.Enabled = true;
                        inserisciSopraTopMenu.Enabled = true;
                    }
                    else
                    {
                        inserisciSopraContextMenu.Enabled = false;
                        inserisciSopraTopMenu.Enabled = false;
                    }

                    if ((DateTime)dataGridParametriD["Inizio Validità", index].Value > DateTime.Today.AddDays(-1))
                    {
                        modificareValoreContextMenu.Enabled = true;
                        cancellaParametroContextMenu.Enabled = true;
                        modificaTopMenu.Enabled = true;
                        elimiaTopMenu.Enabled = true;
                    }
                    else
                    {
                        modificareValoreContextMenu.Enabled = false;
                        cancellaParametroContextMenu.Enabled = false;
                        modificaTopMenu.Enabled = false;
                        elimiaTopMenu.Enabled = false;
                    }
                }
                else
                {
                    modificareValoreContextMenu.Enabled = false;
                    cancellaParametroContextMenu.Enabled = false;
                    modificaTopMenu.Enabled = false;
                    elimiaTopMenu.Enabled = false;
                }
            }
        }

        #region Parametri Giornalieri

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
                }) ?? new DataTable();

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
                    if (fineValidita.Year != 9999)
                        newRow["Fine Validità"] = fineValidita;
                    newRow["Valore"] = decimal.Parse(val["Valore"].ToString(), CultureInfo.CurrentUICulture);

                    valCorretti.Rows.Add(newRow);
                }

                dataGridParametriD.DataSource = valCorretti;
                dataGridParametriD.Columns["Inizio Validità"].DefaultCellStyle.FormatProvider = CultureInfo.InstalledUICulture;
                dataGridParametriD.Columns["Fine Validità"].DefaultCellStyle.FormatProvider = CultureInfo.InstalledUICulture;
                dataGridParametriD.Columns["Fine Validità"].DefaultCellStyle.NullValue = "-";
                dataGridParametriD.Columns["Valore"].DefaultCellStyle.Format = "0.#########";

                foreach (DataGridViewColumn c in dataGridParametriD.Columns)
                    c.SortMode = DataGridViewColumnSortMode.NotSortable;

                if (dataGridParametriD.Rows.Count > 0)
                {
                    dataGridParametriD.CurrentCell = dataGridParametriD["Inizio Validità", dataGridParametriD.Rows.Count - 1];
                }
            }
            else
            {
                dataGridParametriD.DataSource = null;
            }
        }

        private void dataGridParametriD_MouseEnter(object sender, EventArgs e)
        {
            if (!_onEdit)
                dataGridParametriD.Select();
        }
        private void dataGridParametriD_MouseDown(object sender, MouseEventArgs e)
        {
            if (!dataGridParametriD.CurrentCell.IsInEditMode)
            {
                var info = dataGridParametriD.HitTest(e.X, e.Y);
                if (info.RowIndex >= 0 && info.ColumnIndex >= 0)
                    dataGridParametriD.CurrentCell = dataGridParametriD[info.ColumnIndex, info.RowIndex];
            }
            
        }

        private void dataGridParametriD_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (_fromMenuStrip)
            {
                _onEdit = true;

                modificareValoreContextMenu.Enabled = false;
                cancellaParametroContextMenu.Enabled = false;
                modificaTopMenu.Enabled = false;
                elimiaTopMenu.Enabled = false;
                inserisciSopraContextMenu.Enabled = false;
                inserisciSopraTopMenu.Enabled = false;
                inserisciSottoTopMenu.Enabled = false;
                inserisciSottoContextMenu.Enabled = false;

                if (e.ColumnIndex != 0 && e.ColumnIndex != 2)
                    e.Cancel = true;
            }
            else
            {
                e.Cancel = true;
            }                
        }
        private void dataGridParametriD_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (dataGridParametriD.IsCurrentCellDirty && dataGridParametriD.EditingControl != null)
            {
                string value = e.FormattedValue.ToString();

                //salvo la vecchia dataIV per fare l'update
                _currDataIV = (DateTime)dataGridParametriD["Inizio Validità", e.RowIndex].Value;
                _currentCellModified = true;

                switch (e.ColumnIndex)
                {
                    case 0:
                        DateTime date = new DateTime();

                        if (DateTime.TryParseExact(value, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out date) ||
                            DateTime.TryParseExact(value, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out date) ||
                            DateTime.TryParseExact(value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out date) ||
                            DateTime.TryParseExact(value, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out date) ||
                            value == "-")
                        {
                            dataGridParametriD.EditingControl.Text = date.ToString("dd/MM/yyyy");

                            if (date < DateTime.Today)
                            {
                                MessageBox.Show("La data di inizio vaidità non può essere antecedente a oggi!", Simboli.nomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                e.Cancel = true;
                            }
                        }
                        else
                        {
                            e.Cancel = true;
                        }
                        break;
                    case 2:
                        decimal number;
                        if (value == "")
                        {
                            MessageBox.Show("Non è possibile lasciare il campo Valore vuoto!", Simboli.nomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            e.Cancel = true;
                        } 
                        else if (decimal.TryParse(value, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out number) ||
                            decimal.TryParse(value, NumberStyles.AllowDecimalPoint, CultureInfo.InstalledUICulture, out number))
                        {
                            dataGridParametriD.EditingControl.Text = number.ToString(CultureInfo.InstalledUICulture);
                        }
                        else
                        {
                            MessageBox.Show("Il valore inserito non è un numero valido!", Simboli.nomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            e.Cancel = true;
                        }
                        break;
                }
            }
            else
            {
                _currentCellModified = false;
            }
        }
        private void dataGridParametriD_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (!_currentCellModified)
            {
                RefreshMenuItems();
                _onEdit = false;
                _fromMenuStrip = false;
            }
        }
        private void dataGridParametriD_CurrentCellChanged(object sender, EventArgs e)
        {
            RefreshMenuItems();
        }

        private void dataGridParametriD_RowDirtyStateNeeded(object sender, QuestionEventArgs e)
        {
            e.Response = _onEdit;
        }

        private void dataGridParametriD_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {
            DateTime currentIV = (DateTime)dataGridParametriD["Inizio Validità", e.RowIndex].Value;

            //controllo, se esiste, la dataIV della riga successiva ed eventualmente aggiusto la fine validità
            if (e.RowIndex < dataGridParametriD.Rows.Count - 1)
            {
                DateTime subsequentIV = (DateTime)dataGridParametriD["Inizio Validità", e.RowIndex + 1].Value;
                DateTime currentFV = dataGridParametriD["Fine Validità", e.RowIndex].Value is DBNull ? subsequentIV : (DateTime)dataGridParametriD["Fine Validità", e.RowIndex].Value;

                if (currentIV >= subsequentIV)
                {
                    MessageBox.Show("La data di inizio validità della riga corrente va in conflitto con quella della successiva.", Simboli.nomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                    return;
                }

                if (subsequentIV - currentFV != new TimeSpan(1, 0, 0, 0))
                {
                    dataGridParametriD["Fine Validità", e.RowIndex].Value = subsequentIV.AddDays(-1);
                }
            }

            //controllo, se esiste, la dataFV della riga precedente ed eventualmente la aggiorno
            if (e.RowIndex > 0)
            {

                DateTime precedingIV = (DateTime)dataGridParametriD["Inizio Validità", e.RowIndex - 1].Value;
                DateTime precedingFV = (DateTime)dataGridParametriD["Fine Validità", e.RowIndex - 1].Value;

                if (currentIV - precedingIV < new TimeSpan(1, 0, 0, 0))
                {
                    MessageBox.Show("La data di inizio validità della riga corrente va in conflitto con quella della precedente.", Simboli.nomeApplicazione + " - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;

                    return;
                }

                if (currentIV - precedingFV != new TimeSpan(1, 0, 0, 0))
                    dataGridParametriD["Fine Validità", e.RowIndex - 1].Value = currentIV.AddDays(-1);
            }

            //vedo se devo eliminare la data di fine validità
            if (e.RowIndex == dataGridParametriD.Rows.Count - 1 && dataGridParametriD["Fine Validità", e.RowIndex].Value != DBNull.Value)
            {
                dataGridParametriD["Fine Validità", e.RowIndex].Value = DBNull.Value;
            }
        }
        private void dataGridParametriD_RowValidated(object sender, DataGridViewCellEventArgs e)
        {
            DataRowView parRow = (DataRowView)cmbParametriD.SelectedValue;

            if (dataGridParametriD.IsCurrentRowDirty)
            {
                DateTime precIV = (DateTime)dataGridParametriD["Inizio Validità", e.RowIndex - 1].Value;
                DateTime newIV = (DateTime)dataGridParametriD["Inizio Validità", e.RowIndex].Value;
                DateTime newFV = dataGridParametriD["Fine Validità", e.RowIndex].Value is DBNull ? DateTime.MaxValue : (DateTime)dataGridParametriD["Fine Validità", e.RowIndex].Value;
                decimal value = (decimal)dataGridParametriD["Valore", e.RowIndex].Value;

                //insert or update
                if (_isUpdate)
                {
                    Utility.DataBase.Insert(Utility.DataBase.SP.UPDATE_PARAMETRO, new Core.QryParams()
                    {
                        {"@IdEntita", parRow["IdEntita"]},
                        {"@IdTipologiaParametro", parRow["IdParametro"]},
                        {"@CurrDataIV", _currDataIV.ToString("yyyyMMdd")},
                        {"@NewDataIV", newIV.ToString("yyyyMMdd")},
                        {"@NewDataFV", newFV.ToString("yyyyMMdd")},
                        {"@Valore", value},
                        {"@Dettaglio", "D"}
                    });


                }
                else
                {
                    Utility.DataBase.Insert(Utility.DataBase.SP.INSERT_PARAMETRO, new Core.QryParams()
                    {
                        {"@IdEntita", parRow["IdEntita"]},
                        {"@IdTipologiaParametro", parRow["IdParametro"]},
                        {"@PrecDataIV", precIV.ToString("yyyyMMdd")},
                        {"@NewDataIV", newIV.ToString("yyyyMMdd")},
                        {"@NewDataFV", newFV.ToString("yyyyMMdd")},
                        {"@Valore", value},
                        {"@Dettaglio", "D"}
                    });
                }
            }
            RefreshMenuItems();
            _onEdit = false;
            _fromMenuStrip = false;
        }
        
        private void modificareValoreContextMenu_Click(object sender, EventArgs e)
        {
            if (dataGridParametriD.CurrentCell.OwningColumn.Name == "Fine Validità")
                dataGridParametriD.CurrentCell = dataGridParametriD["Inizio Validità", dataGridParametriD.CurrentCell.RowIndex];

            _isUpdate = true;
            _fromMenuStrip = true;
            dataGridParametriD.BeginEdit(false);
        }
        private void modificaTopMenu_Click(object sender, EventArgs e)
        {
            modificareValoreContextMenu_Click(dataGridParametriD, e);
        }
        private void inserisciSopraContextMenu_Click(object sender, EventArgs e)
        {
            _isUpdate = false;

            int index = dataGridParametriD.CurrentRow.Index;

            DataTable dt = (DataTable)dataGridParametriD.DataSource;

            DateTime precedingFV = (DateTime)dt.Rows[index - 1]["Fine Validità"];

            //inserisco la nuova riga
            DataRow r = dt.NewRow();
            r["Inizio Validità"] = precedingFV;
            r["Fine Validità"] = precedingFV;
            dt.Rows.InsertAt(r, index);

            //metto la datagrid in modifica
            dataGridParametriD.CurrentCell = dataGridParametriD["Valore", index];
            _fromMenuStrip = true;
            dataGridParametriD.BeginEdit(false);
        }
        private void inserisciSottoContextMenu_Click(object sender, EventArgs e)
        {
            _isUpdate = false;

            int index = dataGridParametriD.CurrentRow.Index;
            
            DataTable dt = (DataTable)dataGridParametriD.DataSource;

            DateTime precedingFV = dt.Rows[index]["Fine Validità"] is DBNull ? DateTime.Today.AddDays(-1) : (DateTime)dt.Rows[index]["Fine Validità"];
            DateTime iv = precedingFV.AddDays(1);
            DateTime fv = dt.Rows[index]["Fine Validità"] is DBNull ? DateTime.MaxValue : ((DateTime)dt.Rows[index + 1]["Inizio Validità"]).AddDays(-1);

            if (iv > fv)
                iv = fv;

            //inserisco la nuova riga
            DataRow r = dt.NewRow();
            r["Inizio Validità"] = iv;
            if (fv != DateTime.MaxValue)
                r["Fine Validità"] = fv;
            dt.Rows.InsertAt(r, index + 1);

            //metto la datagrid in modifica
            dataGridParametriD.CurrentCell = dataGridParametriD["Valore", index + 1];
            _fromMenuStrip = true;
            dataGridParametriD.BeginEdit(false);
        }
        private void inserisciSopraTopMenu_Click(object sender, EventArgs e)
        {
            inserisciSopraContextMenu_Click(dataGridParametriD, e);
        }
        private void inserisciSottoTopMenu_Click(object sender, EventArgs e)
        {
            inserisciSottoContextMenu_Click(dataGridParametriD, e);
        }

        private void cancellaParametroContextMenu_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Eliminare la riga?", Simboli.nomeApplicazione, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == System.Windows.Forms.DialogResult.Yes)
            {
                DataRowView parRow = (DataRowView)cmbParametriD.SelectedValue;

                DateTime dataIV = (DateTime)dataGridParametriD["Inizio Validità", dataGridParametriD.CurrentRow.Index].Value;
                DateTime dataFV = dataGridParametriD["Fine Validità", dataGridParametriD.CurrentRow.Index].Value is DBNull ? DateTime.MaxValue : (DateTime)dataGridParametriD["Fine Validità", dataGridParametriD.CurrentRow.Index].Value;


                Utility.DataBase.Insert(Utility.DataBase.SP.DELETE_PARAMETRO, new Core.QryParams()
                {
                    {"@IdEntita", parRow["IdEntita"]},
                    {"@IdTipologiaParametro", parRow["IdParametro"]},
                    {"@DataIV", dataIV.ToString("yyyyMMdd")},
                    {"@DataFV", dataFV.ToString("yyyyMMdd")},
                    {"@Dettaglio", "D"}
                });

                int index = cmbParametriD.SelectedIndex;
                cmbParametriD.SelectedIndex = -1;
                cmbParametriD.SelectedIndex = index;
            }
        }
        private void elimiaTopMenu_Click(object sender, EventArgs e)
        {
            cancellaParametroContextMenu_Click(dataGridParametriD, e);
        }
        
        #endregion
    }
}
