using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
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
        double[] _pMin;
        List<object> _sigleRampa;
        int _childWidth;
        int _oreFermata;
        Excel.Worksheet _ws;
        object[] _valoriPQNR;
        DefinedNames _nomiDefiniti;
        string _siglaEntita;
        string _suffissoData;
        Tuple<int, int>[] _profiloPQNR;

        #endregion

        #region Costruttore

        public FormRampe(Excel.Range rng)
        {
            InitializeComponent();
            this.Text = Simboli.nomeApplicazione + " - Rampe";

            if (DataBase.OpenConnection())
            {
                _ws = (Excel.Worksheet)Workbook.WB.ActiveSheet;
                _nomiDefiniti = new DefinedNames(_ws.Name);

                string nome = _nomiDefiniti[rng.Row, rng.Column][0];
                _siglaEntita = nome.Split(Simboli.UNION[0])[0];

                _suffissoData = Regex.Match(nome, @"DATA\d+").Value;
                _suffissoData = _suffissoData == "" ? "DATA1" : _suffissoData;

                DataView proprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPROPRIETA].DefaultView;
                proprieta.RowFilter = "SiglaEntita = '" + _siglaEntita + "' AND SiglaProprieta = 'SISTEMA_COMANDI_PRIF'";
                _pRif = 0;
                if (proprieta.Count > 0)
                    _pRif = Double.Parse(proprieta[0]["Valore"].ToString());

                _entitaRampa = DataBase.LocalDB.Tables[DataBase.Tab.ENTITARAMPA].DefaultView;
                _entitaRampa.RowFilter = "SiglaEntita = '" + _siglaEntita + "'";
                _sigleRampa = _entitaRampa.ToTable(false, "SiglaRampa").AsEnumerable().Select(r => r["SiglaRampa"]).ToList();

                DataView assetti = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAASSETTO].DefaultView;
                assetti.RowFilter = "SiglaEntita = '" + _siglaEntita + "'";

                _profiloPQNR = _nomiDefiniti[DefinedNames.GetName(_siglaEntita, "PQNR_PROFILO", _suffissoData)];
                object[,] values = _ws.Range[_ws.Cells[_profiloPQNR[0].Item1, _profiloPQNR[0].Item2], _ws.Cells[_profiloPQNR[0].Item1, _profiloPQNR[_profiloPQNR.Length - 1].Item2]].Value;
                _valoriPQNR = values.Cast<object>().ToArray();

                //TODO controllare se si può semplificare
                _pMin = new double[_valoriPQNR.Length];
                int numAssetto = 1;
                foreach (DataRowView assetto in assetti)
                {
                    Tuple<int, int>[] cellePmin = _nomiDefiniti[DefinedNames.GetName(_siglaEntita, "PMIN_TERNA_ASSETTO" + numAssetto, _suffissoData)];
                    object[,] pMinOraria = _ws.Range[_ws.Cells[cellePmin[0].Item1, cellePmin[0].Item2], _ws.Cells[cellePmin[0].Item1, cellePmin[cellePmin.Length - 1].Item2]].Value;
                    //object[] pMinOraria = tmppMinOraria.Cast<object>().ToArray();
                    for (int i = 0; i < pMinOraria.GetLength(1); i++)
                    {
                        _pMin[i] = Math.Min(_pMin[i], (double)(pMinOraria[1, i + 1] ?? 0d));
                    }
                    numAssetto++;
                }

                _oreGiorno = _valoriPQNR.Length;
                _oreFermata = int.Parse(DataBase.DB.Select(DataBase.SP.GET_ORE_FERMATA, "@SiglaEntita=" + _siglaEntita).Rows[0]["OreFermata"].ToString());

                _childWidth = panelValoriRampa.Width / _oreGiorno;
                this.Width = tableLayoutDesRampa.Width + (_childWidth * _oreGiorno) + (this.Padding.Left);
                this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;

                DataBase.DB.CloseConnection();
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
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + _siglaEntita + "'";

            lbDesEntita.Text = categoriaEntita[0]["DesEntita"].ToString() + "   -   Potenza rif = " + _pRif + "MW   -   Ore fermata = " + _oreFermata;

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
            //DataTable o = initOutTable();
            object[] intestazione = new object[_oreGiorno];
            object[,] valori = new object[24, _oreGiorno];

            for (int i = 0; i < _oreGiorno; i++)
            {
                _pMin[i] = _pMin[i] < _pRif ? _pRif : _pMin[i];

                var oraX = panelValoriRampa.Controls.OfType<TableLayoutPanel>().FirstOrDefault(r => r.Name == "H" + (i + 1));
                var check = oraX.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked);
                int pos = oraX.Controls.IndexOf(check) - 1;

                intestazione[i] = _sigleRampa[pos];
                _entitaRampa.RowFilter = "SiglaEntita = '" + _siglaEntita + "' AND SiglaRampa = '" + _sigleRampa[pos] + "'";

                for (int j = 0; j < 24; j++)
                {
                    if (_entitaRampa[0]["Q" + (j + 1)] != DBNull.Value)
                    {
                        valori[j, i] = ((int)_entitaRampa[0]["Q" + (j + 1)]) * _pRif / _pMin[i];
                    }
                }
            }

            Excel.Range rng = _ws.Range[_ws.Cells[_profiloPQNR[0].Item1, _profiloPQNR[0].Item2], _ws.Cells[_profiloPQNR[0].Item1, _profiloPQNR[_profiloPQNR.Length - 1].Item2]];
            rng.Value = intestazione;

            Tuple<int,int>[] valoriRampe = _nomiDefiniti.GetByFilter(DefinedNames.Fields.Foglio + " = '" + _ws.Name + "' AND " +
                                                 DefinedNames.Fields.Nome + " LIKE '" + DefinedNames.GetName(_siglaEntita, "PQNR") + "%' AND " +
                                                 DefinedNames.Fields.Nome + " NOT LIKE '%PROFILO%' AND " +
                                                 DefinedNames.Fields.Nome + " LIKE '%" + _suffissoData + "%'");

            _ws.Range[_ws.Cells[valoriRampe[0].Item1, valoriRampe[0].Item2], _ws.Cells[valoriRampe[valoriRampe.Length - 1].Item1, valoriRampe[valoriRampe.Length - 1].Item2]].Value = valori;
            


            Handler.StoreEdit(_ws, rng);
            DataBase.SalvaModificheDB();
        }
        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #endregion
    }
}