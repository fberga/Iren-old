using Iren.ToolsExcel.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Forms
{
    public partial class FormRampe : Form
    {
        #region Variabili

        int _oreGiorno = 24;
        DataView _entitaRampa;
        double _pRif;
        double?[] _pMin;
        string _desEntita;
        List<object> _sigleRampa;
        int _childWidth;
        int _oreFermata;
        Excel.Worksheet _ws;
        object[] _valoriPQNR;
        Tuple<int, int>[] _profiloPQNR;

        #endregion

        #region Costruttore

        public FormRampe(DefinedNames nomiDefiniti, Excel.Range rng)
        {
            InitializeComponent();
            this.Text = Simboli.nomeApplicazione + " - Rampe";

            if (CommonFunctions.DB.OpenConnection())
            {
                _ws = (Excel.Worksheet)CommonFunctions.WB.ActiveSheet;

                string nome = nomiDefiniti[rng.Row, rng.Column][0];
                string up = nome.Split(Simboli.UNION[0])[0];

                string suffissoData = Regex.Match(nome, @"DATA\d+").Value;
                suffissoData = suffissoData == "" ? "DATA1" : suffissoData;

                DataView proprieta = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAPROPRIETA].DefaultView;
                proprieta.RowFilter = "SiglaEntita = '" + up + "' AND SiglaProprieta = 'SISTEMA_COMANDI_PRIF'";
                _pRif = 0;
                if (proprieta.Count > 0)
                    _pRif = Double.Parse(proprieta[0]["Valore"].ToString());

                DataView categoriaEntita = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.CATEGORIAENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + up + "'";
                _desEntita = categoriaEntita[0]["DesEntita"].ToString();

                _entitaRampa = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITARAMPA].DefaultView;
                _entitaRampa.RowFilter = "SiglaEntita = '" + up + "'";
                _sigleRampa = _entitaRampa.ToTable(false, "SiglaRampa").AsEnumerable().Select(r => r["SiglaRampa"]).ToList();

                _profiloPQNR = nomiDefiniti[DefinedNames.GetName(up, "PQNR_PROFILO", suffissoData)];
                object[,] values = _ws.Range[_ws.Cells[_profiloPQNR[0].Item1, _profiloPQNR[0].Item2], _ws.Cells[_profiloPQNR[0].Item1, _profiloPQNR[_profiloPQNR.Length - 1].Item2]].Value;
                _valoriPQNR = values.Cast<object>().ToArray();

                DataView assetti = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAASSETTO].DefaultView;
                assetti.RowFilter = "SiglaEntita = '" + up + "'";

                //TODO controllare se si può semplificare
                _pMin = new double?[_valoriPQNR.Length];
                int numAssetto = 1;
                foreach (DataRowView assetto in assetti)
                {
                    Tuple<int, int>[] cellePmin = nomiDefiniti[DefinedNames.GetName(up, "PMIN_TERNA_ASSETTO" + numAssetto, suffissoData)];
                    object[,] tmppMinOraria = _ws.Range[_ws.Cells[cellePmin[0].Item1, cellePmin[0].Item2], _ws.Cells[cellePmin[0].Item1, cellePmin[cellePmin.Length - 1].Item2]].Value;
                    double?[] pMinOraria = tmppMinOraria.Cast<double?>().ToArray();
                    for (int i = 0; i < pMinOraria.Length; i++)
                    {
                        _pMin[i] = Math.Min(_pMin[i] ?? pMinOraria[i] ?? 0, pMinOraria[i] ?? 0);
                    }
                    numAssetto++;
                }

                _oreGiorno = _valoriPQNR.Length;
                _oreFermata = int.Parse(CommonFunctions.DB.Select("spGetOreFermata", "@SiglaEntita=" + up).Rows[0]["OreFermata"].ToString());

                _childWidth = panelValoriRampa.Width / _oreGiorno;
                this.Width = tableLayoutDesRampa.Width + (_childWidth * _oreGiorno) + (this.Padding.Left);
                this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;

                CommonFunctions.DB.CloseConnection();
            }
        }

        #endregion

        #region Metodi

        private DataTable initOutTable()
        {
            DataTable dt = new DataTable()
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

        #endregion

        #region Eventi

        private void frmRAMPE_Load(object sender, EventArgs e)
        {
            lbDesEntita.Text = _desEntita + "   -   Potenza rif = " + _pRif + "MW   -   Ore fermata = " + _oreFermata;

            tableLayoutDesRampa.Controls.Clear();
            tableLayoutDesRampa.ColumnStyles.Clear();
            tableLayoutDesRampa.RowStyles.Clear();

            tableLayoutRampe.Controls.Clear();
            tableLayoutRampe.ColumnStyles.Clear();
            tableLayoutRampe.RowStyles.Clear();

            tableLayoutRampe.CellPaint += tb_CellPaint;

            tableLayoutDesRampa.RowCount = _entitaRampa.Count + 1;
            tableLayoutRampe.RowCount = _entitaRampa.Count;
            float rowHeightPercentage = 100f / (_entitaRampa.Count + 1) / 100;
            tableLayoutDesRampa.ColumnCount = 2;
            tableLayoutRampe.ColumnCount = _entitaRampa.Table.Columns.Count - 2;

            //scrivo gli header della griglia di visualizzazione delle rampe
            tableLayoutRampe.RowStyles.Add(new RowStyle(SizeType.Percent, rowHeightPercentage));
            for (int i = 0; i < _entitaRampa.Table.Columns.Count - 2; i++)
            {
                switch (i)
                {
                    case 0:
                        tableLayoutRampe.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 95f));
                        tableLayoutRampe.Controls.Add(new Label() { Text = "Rampa", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = new Font(this.Font, FontStyle.Bold), BackColor = System.Drawing.Color.Transparent }, i, 0);
                        break;
                    case 1:
                        tableLayoutRampe.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 75f));
                        tableLayoutRampe.Controls.Add(new Label() { Text = "Fermo da", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font(this.Font, FontStyle.Bold), BackColor = System.Drawing.Color.Transparent }, i, 0);
                        break;
                    case 2:
                        tableLayoutRampe.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 75f));
                        tableLayoutRampe.Controls.Add(new Label() { Text = "Fermo a", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font(this.Font, FontStyle.Bold), BackColor = System.Drawing.Color.Transparent }, i, 0);
                        break;
                    default:
                        tableLayoutRampe.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (tableLayoutRampe.Width - 245f) / (tableLayoutRampe.ColumnCount - 2)));
                        tableLayoutRampe.Controls.Add(new Label() { Text = _entitaRampa.Table.Columns[i + 2].ColumnName, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font(this.Font, FontStyle.Bold), BackColor = System.Drawing.Color.Transparent }, i, 0);
                        break;
                }
            }

            int y = 1;
            tableLayoutDesRampa.RowStyles.Add(new RowStyle(SizeType.Percent, rowHeightPercentage));
            tableLayoutDesRampa.Controls.Add(new Label() { Text = "Tutte", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font(this.Font, FontStyle.Bold), BackColor = System.Drawing.Color.Transparent}, 3, 0);

            tableLayoutDesRampa.CellPaint += tb_CellPaint;

            foreach (DataRowView rampa in _entitaRampa)
            {
                tableLayoutDesRampa.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 0.65f));
                tableLayoutDesRampa.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 0.25f));

                tableLayoutDesRampa.RowStyles.Add(new RowStyle(SizeType.Percent, rowHeightPercentage));

                tableLayoutDesRampa.Controls.Add(new Label() { Text = rampa["DesRampa"].ToString(), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, BackColor = System.Drawing.Color.Transparent, Font = new Font(this.Font, FontStyle.Bold) }, 0, y);

                RadioButton rb = new RadioButton() { Name = rampa["SiglaRampa"].ToString(), Dock = DockStyle.Fill, CheckAlign = ContentAlignment.MiddleCenter, BackColor = System.Drawing.Color.Transparent };
                rb.CheckedChanged += rbTutti_CheckedChanged;

                tableLayoutDesRampa.Controls.Add(rb, 1, y);

                tableLayoutRampe.RowStyles.Add(new RowStyle(SizeType.Percent, rowHeightPercentage));
                for (int i = 0; i < _entitaRampa.Table.Columns.Count - 2; i++)
                {
                    switch (i)
                    {
                        case 0:
                            tableLayoutRampe.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105f));
                            tableLayoutRampe.Controls.Add(new Label() { Text = rampa["DesRampa"].ToString(), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, BackColor = System.Drawing.Color.Transparent}, i, y);
                            break;
                        case 1:
                            tableLayoutRampe.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
                            tableLayoutRampe.Controls.Add(new Label() { Text = rampa["FermoDa"].ToString(), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, BackColor = System.Drawing.Color.Transparent}, i, y);
                            break;
                        case 2:
                            tableLayoutRampe.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
                            tableLayoutRampe.Controls.Add(new Label() { Text = rampa["FermoA"].ToString(), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, BackColor = System.Drawing.Color.Transparent}, i, y);
                            break;
                        default:
                            tableLayoutRampe.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (tableLayoutRampe.Width - 225f) / (tableLayoutRampe.ColumnCount - 2)));
                            tableLayoutRampe.Controls.Add(new Label() { Text = rampa["Q" + (i - 2)].ToString(), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, BackColor = System.Drawing.Color.Transparent}, i, y);
                            break;
                    }
                }
                y++;

            }

            int left = 2;            

            for (int i = 1; i <= _oreGiorno; i++)
            {
                TableLayoutPanel tb = new TableLayoutPanel()
                {
                    Name = "H" + i,
                    ColumnCount = 1,
                    RowCount = _entitaRampa.Count + 1,
                    Height = panelValoriRampa.Height,
                    Width = _childWidth,
                    Left = left - 1,
                    CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                };
                tb.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, _childWidth));
                tb.RowStyles.Add(new RowStyle(SizeType.Percent, rowHeightPercentage));
                tb.Controls.Add(new Label() { Text = "H" + i, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font(this.Font, FontStyle.Bold), BackColor = System.Drawing.Color.Transparent }, 0, 0);
                
                tb.CellPaint += tb_CellPaint;

                y = 1;
                foreach (DataRowView rampa in _entitaRampa)
                {
                    tb.RowStyles.Add(new RowStyle(SizeType.Percent, rowHeightPercentage));

                    RadioButton rb = new RadioButton() { Dock = DockStyle.Fill, CheckAlign = ContentAlignment.MiddleCenter, BackColor = System.Drawing.Color.Transparent};
                    rb.CheckedChanged += rbOre_CheckedChanged;

                    tb.Controls.Add(rb, 0, y++);
                }
                left = tb.Right;
                panelValoriRampa.Controls.Add(tb);
            }

            //carico valori PQNR
            if (_valoriPQNR[0] != null)
            {
                for (int i = 0; i < _valoriPQNR.Length; i++)
                {
                    ((RadioButton)Controls.Find("H" + (i + 1), true)[0].Controls[_sigleRampa.IndexOf(_valoriPQNR[i]) + 1]).Checked = true;
                }
            }
            else
            {
                tableLayoutDesRampa.Controls.OfType<RadioButton>().First().Checked = true;
            }
        }
        private void tb_CellPaint(object sender, TableLayoutCellPaintEventArgs e)
        {
            if (((TableLayoutPanel)sender).Name == "tableLayoutRampe")
            {
                if (e.Row == 0)
                {
                    e.Graphics.FillRectangle(Brushes.Gray, e.CellBounds);
                }
                else
                {
                    e.Graphics.FillRectangle(Brushes.Gray, e.CellBounds);
                }
            }
            else
            {
                if (((TableLayoutPanel)sender).Name == "tableLayoutDesRampa")
                {
                    if (e.Column > 0 & e.Row >= 0)
                    {
                        e.Graphics.FillRectangle(Brushes.LightGreen, e.CellBounds);
                    }
                    else
                    {
                        if (e.Column == 0 & e.Row != 0)
                        {
                            e.Graphics.FillRectangle(Brushes.LightGreen, e.CellBounds);
                        }
                    }
                }
                else
                {
                    if (e.Row == 0)
                    {
                        e.Graphics.FillRectangle(Brushes.LightGreen, e.CellBounds);
                    }
                    else
                    {
                        e.Graphics.FillRectangle(Brushes.LightGray, e.CellBounds);
                    }
                }
            }
        }
        private void rbOre_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton rb = (RadioButton)sender;
            int pos = rb.Parent.Controls.GetChildIndex(rb);
            bool allChecked = true;
            for (int i = 1; i <= _oreGiorno; i++)
            {
                RadioButton rb1 = (RadioButton)Controls.Find("H" + i, true)[0].Controls[pos];
                allChecked = allChecked & rb1.Checked;
            }

            ((RadioButton)Controls.Find(_sigleRampa[pos - 1].ToString(), true)[0]).Checked = allChecked;
        }
        private void rbTutti_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton rb = (RadioButton)sender;
            if (rb.Checked)
            {
                int pos = _sigleRampa.IndexOf(rb.Name);
                for (int i = 1; i <= _oreGiorno; i++)
                {
                    RadioButton rb1 = (RadioButton)Controls.Find("H" + i, true)[0].Controls[pos + 1];
                    rb1.Checked = true;
                }
            }
        }
        private void btnApplica_Click(object sender, EventArgs e)
        {
            DataTable o = initOutTable();
            for (int i = 1; i <= _oreGiorno; i++)
            {
                DataRow riga = o.NewRow();

                var oraX = panelValoriRampa.Controls.OfType<TableLayoutPanel>().FirstOrDefault(r => r.Name == "H" + i);
                var check = oraX.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked);
                int pos = oraX.Controls.IndexOf(check) - 1;

                riga["SiglaRampa"] = _sigleRampa[pos];
                _entitaRampa.RowFilter += " AND SiglaRampa = '" + _sigleRampa[pos] + "'";

                for (int j = 1; j <= 24; j++)
                {
                    if (_entitaRampa[0]["Q" + j] != DBNull.Value)
                    {
                        _pMin[i - 1] = _pMin[i - 1] < _pRif ? _pRif : _pMin[i - 1];
                        riga["Q" + j] = ((int)_entitaRampa[0]["Q" + j]) * _pRif / _pMin[i - 1];
                    }
                }

                _entitaRampa.RowFilter = _entitaRampa.RowFilter.Replace(" AND SiglaRampa = '" + _sigleRampa[pos] + "'", "");

                o.Rows.Add(riga);
            }
            Excel.Range rng = _ws.Range[_ws.Cells[_profiloPQNR[0].Item1, _profiloPQNR[0].Item2], _ws.Cells[_profiloPQNR[0].Item1, _profiloPQNR[_profiloPQNR.Length - 1].Item2]];
            rng.Value = o.AsEnumerable().Select(r => r["SiglaRampa"]).ToArray();

            BaseHandler.StoreEdit(_ws, rng);
            CommonFunctions.SalvaModificheDB();
        }
        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #endregion
    }
}