﻿using Iren.ToolsExcel.Base;
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
        object[] _profiloPQNR;
        DefinedNames _definedNames;
        string _siglaEntita;
        string _suffissoData;

        #endregion

        #region Costruttore

        public FormRampe(Excel.Range rng)
        {
            InitializeComponent();
            this.Text = Simboli.nomeApplicazione + " - Rampe";

            if (DataBase.OpenConnection())
            {
                _ws = (Excel.Worksheet)Workbook.WB.ActiveSheet;
                _definedNames = new DefinedNames(_ws.Name, DefinedNames.InitType.NamingOnly);

                string nome = _definedNames.GetNameByAddress(rng.Row, rng.Column);
                _siglaEntita = nome.Split(Simboli.UNION[0])[0];
                _suffissoData = Regex.Match(nome, @"DATA\d+").Value;
                _oreGiorno = Date.GetOreGiorno(_suffissoData);

                _pRif =
                    (from r in DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA].AsEnumerable()
                     where r["SiglaEntita"].Equals(_siglaEntita)
                        && r["SiglaProprieta"].Equals("SISTEMA_COMANDI_PRIF")
                     select Double.Parse(r["Valore"].ToString())).FirstOrDefault();

                _entitaRampa = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_RAMPA].DefaultView;
                _entitaRampa.RowFilter = "SiglaEntita = '" + _siglaEntita + "'";
                _sigleRampa = 
                    (from r in _entitaRampa.ToTable().AsEnumerable()
                     select r["SiglaRampa"]).ToList();

                int assetti = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_ASSETTO].AsEnumerable().Count(r => r["SiglaEntita"].Equals(_siglaEntita));

                Range profilo = _definedNames.Get(_siglaEntita, "PQNR_PROFILO", _suffissoData).Extend(colOffset: _oreGiorno);

                object[,] values = _ws.Range[profilo.ToString()].Value;
                _profiloPQNR = values.Cast<object>().ToArray();

                _pMin = new double[_profiloPQNR.Length];
                for (int i = 0; i < _pMin.Length; i++ ) _pMin[i] = double.MaxValue;

                for(int i = 0; i < assetti; i++)
                {
                    Range rngPmin = _definedNames.Get(_siglaEntita, "PMIN_TERNA_ASSETTO" + (i + 1), _suffissoData).Extend(colOffset: _oreGiorno);
                    for (int j = 0; j < _oreGiorno; j++)
                        _pMin[j] = Math.Min(_pMin[j], (double)(_ws.Range[rngPmin.Columns[j].ToString()].Value ?? 0d));
                }

                _oreFermata = int.Parse(DataBase.Select(DataBase.SP.GET_ORE_FERMATA, "@SiglaEntita=" + _siglaEntita).Rows[0]["OreFermata"].ToString());

                _childWidth = panelValoriRampa.Width / _oreGiorno;
                this.Width = tableLayoutDesRampa.Width + (_childWidth * _oreGiorno) + (this.Padding.Left);
                this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            }
        }

        #endregion

        #region Eventi

        public void frmRAMPE_Load(object sender, EventArgs e)
        {
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
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
            if (_profiloPQNR[0] != null)
            {
                for (int i = 0; i < _profiloPQNR.Length; i++)
                {
                    ((RadioButton)Controls.Find("H" + (i + 1), true)[0].Controls[_sigleRampa.IndexOf(_profiloPQNR[i]) + 1]).Checked = true;
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
        public void btnApplica_Click(object sender, EventArgs e)
        {
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
                        valori[j, i] = Math.Round(((int)_entitaRampa[0]["Q" + (j + 1)]) * _pRif / _pMin[i]);
                    }
                }
            }

            Range rngPQNR = _definedNames.Get(_siglaEntita, "PQNR_PROFILO", _suffissoData).Extend(colOffset: Date.GetOreGiorno(_suffissoData));
            _ws.Range[rngPQNR.ToString()].Value = intestazione;

            Range rngPQNRVal = _definedNames.Get(_siglaEntita, "PQNR1", _suffissoData).Extend(rowOffset: 24, colOffset: Date.GetOreGiorno(_suffissoData));
            _ws.Range[rngPQNRVal.ToString()].Value = valori;
            
            Handler.StoreEdit(_ws, _ws.Range[rngPQNR.ToString()]);
            DataBase.SalvaModificheDB();
        }
        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #endregion
    }
}