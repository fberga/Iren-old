using System;
using System.Windows.Forms;
using Iren.FrontOffice.Core;

namespace Iren.FrontOffice.Forms
{
    using DataTable = System.Data.DataTable;
    using DataView = System.Data.DataView;
    using DataRow = System.Data.DataRow;
    using DataColumn = System.Data.DataColumn;
    using DataSet = System.Data.DataSet;
    using DataRowView = System.Data.DataRowView;
    using Workbook = Microsoft.Office.Interop.Excel.Workbook;
    using System.Drawing;

    public partial class frmRAMPE : Form
    {
        DataBase _db;
        Workbook _wb;
        int _numOre = 24;
        DataView _dvRampe;
        DataView _dvCE;
        DataView _dvModRampe;

        public DataSet _out = new DataSet();

        public frmRAMPE(Workbook wb, ref DataBase db)
        {
            InitializeComponent();

            _wb = wb;
            _db = db;
        }

        private void frmRAMPE_Load(object sender, EventArgs e)
        {
            /*DataBase._livelloLOG++;
            DataBase._Log("frmRAMPE_Load");
            
            Function f = new Function(_wb, ref _db);
            Repository r = new Repository(_wb, ref _db);
            _numOre = f.GetOreData(Tools.GetParametroApplicazione<string>(FOConst.PARAMETERS.DATA_ATTIVA, _wb));

            _dvCE = r.GetCategoriaEntita("IREN_60T", FOConst.PARAMETERS.SP_ALL).DefaultView;
            cmbEntita.DataSource = _dvCE;
            cmbEntita.DisplayMember = "DesEntita";

            _dvRampe = r.GetEntitaRampa(FOConst.PARAMETERS.SP_ALL).DefaultView;

            //personalizzazione datagridview con riassunto rampe per entità
            initDGVisualizzaRampe();
            //creazione datagridview di modifica delle rampe
            initDGModificaRampa();

            //lancia l'evento di selezione sulla combobox
            cmbEntita_SelectedIndexChanged(cmbEntita, new EventArgs());*/
        }

        private void initDGModificaRampa()
        {
            DataTable dt = new DataTable() 
            {
                Columns = 
                {
                    {"SiglaEntita", typeof(string)},
                    {"SiglaRampa", typeof(string)},
                    {"DesRampa", typeof(string)},
                    {"Tutti", typeof(bool)}
                }
            };

            for (int i = 1; i <= _numOre; i++)
            {
                dt.Columns.Add("H" + i, typeof(bool));
            }
            foreach (DataRowView rv in _dvRampe)
            {
                DataRow row = dt.NewRow();
                row["SiglaEntita"] = rv["SiglaEntita"];
                row["SiglaRampa"] = rv["SiglaRampa"];
                row["DesRampa"] = rv["DesRampa"];
                row["Tutti"] = false;
                for (int i = 1; i <= _numOre; i++)
                {
                    row["H" + i] = false;
                }
                dt.Rows.Add(row);
            }
            _dvModRampe = dt.DefaultView;

            dgModificaRampa.DataSource = _dvModRampe;
            dgModificaRampa.MultiSelect = false;
            dgModificaRampa.RowHeadersVisible = false;
            dgModificaRampa.AllowUserToAddRows = false;
            dgModificaRampa.AllowUserToDeleteRows = false;
            dgModificaRampa.AllowUserToResizeColumns = false;
            dgModificaRampa.AllowUserToResizeRows = false;
            dgModificaRampa.AllowUserToOrderColumns = false;

            dgModificaRampa.Columns[0].Visible = false;
            dgModificaRampa.Columns[1].Visible = false;

            dgModificaRampa.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgModificaRampa.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgModificaRampa.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgModificaRampa.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgModificaRampa.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        }

        private void initDGVisualizzaRampe()
        {
            dgVisualizzaRampe.DataSource = _dvRampe;
            dgVisualizzaRampe.AllowUserToAddRows = false;
            dgVisualizzaRampe.AllowUserToDeleteRows = false;
            dgVisualizzaRampe.AllowUserToResizeColumns = false;
            dgVisualizzaRampe.AllowUserToResizeRows = false;
            dgVisualizzaRampe.AllowUserToOrderColumns = false;
            dgVisualizzaRampe.ReadOnly = true;
            dgVisualizzaRampe.MultiSelect = false;
            dgVisualizzaRampe.RowHeadersVisible = false;

            dgVisualizzaRampe.Columns[0].Visible = false;
            dgVisualizzaRampe.Columns[1].Visible = false;


            dgVisualizzaRampe.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgVisualizzaRampe.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgVisualizzaRampe.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgVisualizzaRampe.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgVisualizzaRampe.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgVisualizzaRampe.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void cmbEntita_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataRowView vrow = (DataRowView)cmbEntita.SelectedItem;
            _dvRampe.RowFilter = "SiglaEntita = '" + vrow.Row["SiglaEntita"] + "'";
            _dvModRampe.RowFilter = "SiglaEntita = '" + vrow.Row["SiglaEntita"] + "'";
            int count = 0;
            foreach(DataRowView rv in _dvModRampe) 
            {
                if((bool)rv["Tutti"])
                    count++;
            }
            if (count == 0)
                dgModificaRampa["Tutti", dgModificaRampa.FirstDisplayedCell.RowIndex].Value = true;
        }

        private void dgModificaRampa_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            Brush b = new SolidBrush(dgVisualizzaRampe.DefaultCellStyle.BackColor);
            if (e.RowIndex == -1)
            {
                b = new SolidBrush(dgVisualizzaRampe.ColumnHeadersDefaultCellStyle.BackColor);
            }
            else
            {
                if (e.ColumnIndex <= 3)
                {
                    b = new SolidBrush(System.Drawing.Color.Gainsboro);
                }
            }
            e.Graphics.FillRectangle(b, e.CellBounds);

            Pen col = new Pen(Brushes.Black, 2f);
            Pen allCol = new Pen(Brushes.DarkGray, 0.1f);
            switch (e.ColumnIndex)
            {
                case 2:
                    e.Graphics.DrawLine(col,
                        new Point(e.CellBounds.Right, e.CellBounds.Top),
                        new Point(e.CellBounds.Right, e.CellBounds.Bottom));
                    break;
                case 3:
                    e.Graphics.DrawLine(col,
                        new Point(e.CellBounds.Right, e.CellBounds.Top),
                        new Point(e.CellBounds.Right, e.CellBounds.Bottom));
                    break;
                default:
                    e.Graphics.DrawLine(allCol,
                        new Point(e.CellBounds.Right - 1, e.CellBounds.Top - 1),
                        new Point(e.CellBounds.Right - 1, e.CellBounds.Bottom - 1));
                    break;
            }
            e.Graphics.DrawLine(allCol,
                        new Point(0, e.CellBounds.Bottom - 1),
                        new Point(e.CellBounds.Right, e.CellBounds.Bottom - 1));
            e.PaintContent(e.ClipBounds);
            e.Handled = true;
        }

        private void dgVisualizzaRampe_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            Brush b = new SolidBrush(dgVisualizzaRampe.DefaultCellStyle.BackColor);
            if (e.RowIndex == -1)
            {
                b = new SolidBrush(dgVisualizzaRampe.ColumnHeadersDefaultCellStyle.BackColor);
            }
            else
            {
                if (e.ColumnIndex <= 4)
                {
                    b = new SolidBrush(System.Drawing.Color.Gainsboro);
                }
            }
            e.Graphics.FillRectangle(b, e.CellBounds);

            Pen col4 = new Pen(Brushes.Black, 2f);
            Pen allCol = new Pen(Brushes.DarkGray, 0.1f);
            switch (e.ColumnIndex)
            {
                case 4:
                    e.Graphics.DrawLine(col4,
                        new Point(e.CellBounds.Right, e.CellBounds.Top),
                        new Point(e.CellBounds.Right, e.CellBounds.Bottom));
                    break;
                default:
                    e.Graphics.DrawLine(allCol,
                        new Point(e.CellBounds.Right - 1, e.CellBounds.Top - 1),
                        new Point(e.CellBounds.Right - 1, e.CellBounds.Bottom - 1));
                    break;
            }
            e.Graphics.DrawLine(allCol,
                        new Point(0, e.CellBounds.Bottom-1),
                        new Point(e.CellBounds.Right, e.CellBounds.Bottom-1));
            e.PaintContent(e.ClipBounds);
            e.Handled = true;
        }

        private void dgModificaRampa_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            e.Column.SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        private void dgVisualizzaRampe_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            e.Column.SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        private void dgModificaRampa_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //se sto lavorando sulle celle da Tutti in su (quindi sulle ore)
            if (e.ColumnIndex >= dgModificaRampa.Columns["Tutti"].Index)
            {
                //controllo se nella colonna ci sono più valori a true
                int count = 0;
                foreach (DataRowView rv in _dvModRampe)
                    count += (bool)rv[e.ColumnIndex] ? 1 : 0;

                //se ci sono più celle selezionate su una colonna le disabilito
                //utilizzando il DataView non scateno l'evento on change delle celle
                if (count > 1)
                    for(int i = 0; i < _dvModRampe.Count; i++)
                        _dvModRampe[i][e.ColumnIndex] = i == e.RowIndex;


                //controllo se ho selezionato la colonna tutti
                if (e.ColumnIndex == dgModificaRampa.Columns["Tutti"].Index)
                {
                    for (int i = 1; i <= _numOre; i++)
                        dgModificaRampa[e.ColumnIndex + i, e.RowIndex].Value = dgModificaRampa[e.ColumnIndex, e.RowIndex].Value;
                }
            }

            //applico le modifice e setto a conclusa l'operazione sulla cella
            dgModificaRampa.Invalidate();
            dgModificaRampa.EndEdit();
        }

        private void dgModificaRampa_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            //se faccio una modifica al checkbox, forzo il commit così viene richiamato CellValueChanged
            if (dgModificaRampa.IsCurrentCellDirty)
            {
                dgModificaRampa.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dgModificaRampa_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            //se sto lavorando su una cella già true, fermo l'operazione
            if ((bool)dgModificaRampa[e.ColumnIndex, e.RowIndex].Value)
                e.Cancel = true;
        }

        private DataTable initOutTable(string name) 
        {
            DataTable dt = new DataTable(name)
            {
                Columns =
                {
                    {"SiglaRampa", typeof(string)},
                    {"Q1", typeof(Int32)},
                    {"Q2", typeof(Int32)},
                    {"Q3", typeof(Int32)},
                    {"Q4", typeof(Int32)},
                    {"Q5", typeof(Int32)},
                    {"Q6", typeof(Int32)},
                    {"Q7", typeof(Int32)},
                    {"Q8", typeof(Int32)},
                    {"Q9", typeof(Int32)},
                    {"Q10", typeof(Int32)},
                    {"Q11", typeof(Int32)},
                    {"Q12", typeof(Int32)},
                    {"Q13", typeof(Int32)},
                    {"Q14", typeof(Int32)},
                    {"Q15", typeof(Int32)},
                    {"Q16", typeof(Int32)},
                    {"Q17", typeof(Int32)},
                    {"Q18", typeof(Int32)},
                    {"Q19", typeof(Int32)},
                    {"Q20", typeof(Int32)},
                    {"Q21", typeof(Int32)},
                    {"Q22", typeof(Int32)},
                    {"Q23", typeof(Int32)},
                    {"Q24", typeof(Int32)}
                }
            };
            return dt;
        }

        private void btnApplica_Click(object sender, EventArgs e)
        {
            DataRowView vrow = (DataRowView)cmbEntita.SelectedItem;
            DataTable dt;
            if (_out.Tables.Contains(vrow["SiglaEntita"].ToString()))
            {
                if (MessageBox.Show("Esiste già una configurazione per " + vrow["SiglaEntita"] + ".\nLa configurazione verrà sovrascritta. Continuare?", "Attenzione", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                    _out.Tables.Remove(vrow["SiglaEntita"].ToString());
                else
                    return;
            }
            //la tabella contiene sigla rampa + valori rampa per ogni ora del giorno (record 0 -> _numOre - 1)
            dt = initOutTable(vrow["SiglaEntita"].ToString());
            
            for (int i = 1; i <= _numOre; i++)
            {
                DataRow r = dt.NewRow();
                foreach (DataRowView rv in _dvModRampe)
                {
                    if ((bool)rv["H" + i])
                    {
                        r["SiglaRampa"] = rv["SiglaRampa"];

                        DataRow[] rows = _dvRampe.Table.Select("SiglaEntita = '" + vrow["SiglaEntita"] + "' AND SiglaRampa = '" + rv["SiglaRampa"] + "'");

                        if (rows.Length > 0)
                            for (int j = 1; j <= 24; j++)
                                r["Q" + j] = rows[0]["Q" + j];

                        break;
                    }
                }
                dt.Rows.Add(r);
            }

            _out.Tables.Add(dt);
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
