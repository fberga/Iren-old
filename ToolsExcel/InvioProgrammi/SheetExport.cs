using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    class SheetExport : Base.ASheet
    {
        private static Dictionary<string, DateTime> _dataCaricaStruttura = new Dictionary<string, DateTime>();


        protected Excel.Worksheet _ws;
        protected DefinedNames _definedNames;
        protected int _rigaAttiva;
        protected string _mercato;

        public SheetExport(Excel.Worksheet ws)
        {
            _ws = ws;
            _mercato = ws.Name;

            AggiornaParametriSheet();

            _definedNames = new DefinedNames(_ws.Name);
        }

        protected void AggiornaParametriSheet()
        {
            DataView paramApplicazione = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE].DefaultView;

            _struttura = new Struct();

            _struttura.rigaBlock = (int)paramApplicazione[0]["RowBlocco"];
            _struttura.rigaGoto = (int)paramApplicazione[0]["RowGoto"];
            _struttura.colBlock = 2;

        }

        private void Clear()
        {            
            if (_ws.ChartObjects().Count > 0)
                _ws.ChartObjects().Delete();

            _ws.Rows.Delete();
            _ws.Rows.FormatConditions.Delete();
            _ws.Rows.EntireRow.Hidden = false;
            _ws.Rows.Font.Size = 10;
            _ws.Rows.NumberFormat = "General";
            _ws.Rows.Font.Name = "Verdana";
            _ws.Rows.RowHeight = Struct.cell.height.normal;

            _ws.Columns.ColumnWidth = Struct.cell.width.dato;

            _ws.Rows["1:" + (_struttura.rigaBlock - 1)].RowHeight = Struct.cell.height.empty;

            _ws.Rows[_struttura.rigaGoto].RowHeight = Struct.cell.height.normal;

            _ws.Columns[1].ColumnWidth = Struct.cell.width.empty;

            ((Excel._Worksheet)_ws).Activate();
            _ws.Application.ActiveWindow.FreezePanes = false;
            _ws.Cells[_struttura.rigaBlock, 1].Select();
            _ws.Application.ActiveWindow.ScrollColumn = 1;
            _ws.Application.ActiveWindow.ScrollRow = 1;
            _ws.Application.ActiveWindow.FreezePanes = true;
            Workbook.Main.Select();
            _ws.Application.ScreenUpdating = false;
        }
        protected void InitBarraNavigazione()
        {
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;

            Excel.Range gotoBar = _ws.Range[_ws.Cells[2, 2], _ws.Cells[_struttura.rigaGoto + 1, categoriaEntita.Count + 3]];
            gotoBar.Style = "gotoBarStyle";
            gotoBar.BorderAround2(Weight: Excel.XlBorderWeight.xlMedium, Color: 1);

            int i = 3;
            foreach (DataRowView entita in categoriaEntita)
            {
                Excel.Range rng = _ws.Cells[_struttura.rigaGoto, i++];
                rng.Value = entita["DesEntitaBreve"];
                rng.Style = "navBarStyleHorizontal";
            }
        }
        private void InitColumns()
        {
            //definisco tutte le colonne
            DataTable categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA];

            //Calcolo il massimo numero di entità da mettere affiancate
            int maxElementCount =
                (from r in categoriaEntita.AsEnumerable()
                 where r["Gerarchia"] != DBNull.Value
                 group r by r["Gerarchia"] into g
                 select g.Count()).Max();

            int colonnaAttiva = _struttura.colBlock;
            for (int i = 0; i < maxElementCount; i++)
            {
                colonnaAttiva++;
                for (int j = 0; j < 4; j++)
                    _definedNames.AddCol(colonnaAttiva++, "RIF" + (i + 1), "PROGRAMMAQ" + (j + 1));
            }
        }

        public override void LoadStructure()
        {
            SplashScreen.UpdateStatus("Creo struttura " + _mercato);

            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "Gerarchia = '' OR Gerarchia IS NULL";

            if (DataCaricamentoStruttura != DataBase.DataAttiva)
            {
                Clear();
                InitBarraNavigazione();
            }

            InitColumns();

            _rigaAttiva = _struttura.rigaBlock + 1;

            foreach (DataRowView entita in categoriaEntita)
                InitBloccoEntita(entita);

            _definedNames.DumpToDataSet();

            if (DataCaricamentoStruttura != DataBase.DataAttiva || Simboli.Mercato == _ws.Name)
                CaricaInformazioni(all: true);

            if (_dataCaricaStruttura.ContainsKey(_ws.Name))
                _dataCaricaStruttura[_ws.Name] = DataBase.DataAttiva;
            else
                _dataCaricaStruttura.Add(_ws.Name, DataBase.DataAttiva);
        }

        protected void InitBloccoEntita(DataRowView entita)
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
            CreaNomiCelle(entita["SiglaEntita"]);
            
            if (DataCaricamentoStruttura != DataBase.DataAttiva)
            {
                FormattaBloccoEntita(entita["SiglaEntita"], entita["DesEntita"], entita["CodiceRUP"]);
            }

        }
        protected void CreaNomiCelle(object siglaEntita)
        {
            _definedNames.AddName(_rigaAttiva, siglaEntita, "T");
            _rigaAttiva += 2;
            _definedNames.AddName(_rigaAttiva, siglaEntita, "DATA");
            _rigaAttiva += 2;
            _definedNames.AddName(_rigaAttiva, siglaEntita, "UM", "T");
            _rigaAttiva += Date.GetOreGiorno(DataBase.DataAttiva) + 5;
        }
        protected void FormattaBloccoEntita(object siglaEntita, object desEntita, object codiceRUP)
        {
            //Titolo
            Range rng = new Range(_definedNames.GetRowByName(siglaEntita, "T"), _struttura.colBlock, 1, 10);
            Style.RangeStyle(_ws.Range[rng.ToString()], fontSize: 12, merge: true, bold: true, align: Excel.XlHAlign.xlHAlignCenter, borders: "[top:medium,right:medium,bottom:medium,left:medium]");
            _ws.Range[rng.ToString()].Value = "PROGRAMMA A 15 MINUTI " + desEntita;
            _ws.Range[rng.ToString()].RowHeight = 25;

            //Data
            rng = new Range(_definedNames.GetRowByName(siglaEntita, "DATA"), _struttura.colBlock, 1, 5);
            Style.RangeStyle(_ws.Range[rng.ToString()], fontSize: 10, bold: true, align: Excel.XlHAlign.xlHAlignCenter, borders: "[top:medium,right:medium,bottom:medium,left:medium,insidev:medium]", numberFormat: "dd/MM/yyyy");
            _ws.Range[rng.ToString()].RowHeight = 18;
            _ws.Range[rng.Columns[0].ToString()].Value = "Data";
            _ws.Range[rng.Columns[1, 3].ToString()].Merge();
            _ws.Range[rng.Columns[1].ToString()].Value = DataBase.DataAttiva;
            _ws.Range[rng.Columns[4].ToString()].Value = _mercato;

            //Tabella
            DataTable categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA];
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;

            List<DataRow> entitaRif =
                (from r in categoriaEntita.AsEnumerable()
                 where r["Gerarchia"].Equals(siglaEntita)
                 select r).ToList();
            
            bool hasEntitaRif = entitaRif.Count > 0;
            int numEntita = Math.Max(entitaRif.Count, 1);

            rng = new Range(_definedNames.GetRowByName(siglaEntita, "UM", "T"), _struttura.colBlock, 1, 5 * numEntita);
            for (int i = 0; i < numEntita; i++)
            {
                informazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND Visibile = '1' " + (hasEntitaRif ? "AND SiglaEntitaRif = '" + entitaRif[i]["SiglaEntita"] + "'" : "");
                
                //range grande come tutta la tabella
                rng = new Range(_definedNames.GetRowByName(siglaEntita, "UM", "T"), _definedNames.GetColFromName("RIF" + (i + 1), "PROGRAMMAQ1") - 1, Date.GetOreGiorno(DataBase.DataAttiva) + 2, 5);

                Style.RangeStyle(_ws.Range[rng.ToString()], borders: "[top:medium,right:medium,bottom:medium,left:medium,insideH:thin,insideV:thin]", align: Excel.XlHAlign.xlHAlignCenter, numberFormat: informazioni[0]["Formato"]);
                Style.RangeStyle(_ws.Range[rng.Rows[1, rng.Rows.Count - 1].Columns[0].ToString()], backColor: 15, bold: true, align: Excel.XlHAlign.xlHAlignLeft);
                Style.RangeStyle(_ws.Range[rng.Rows[0].ToString()], backColor: 15, bold: true, fontSize: 11);
                Style.RangeStyle(_ws.Range[rng.Rows[1].ToString()], backColor: 15, bold: true);
                _ws.Range[rng.Rows[0].Columns[1, rng.Columns.Count - 1].ToString()].Merge();
                if (hasEntitaRif)
                    _ws.Range[rng.Rows[0].ToString()].Value = new object[] { "UM", entitaRif[i]["CodiceRUP"] is DBNull ? entitaRif[i]["DesEntita"] : entitaRif[i]["CodiceRUP"] };
                else
                    _ws.Range[rng.Rows[0].ToString()].Value = new object[] { "UM", codiceRUP is DBNull ? desEntita : codiceRUP };

                for (int h = 1; h <= Date.GetOreGiorno(DataBase.DataAttiva); h++)
                    _ws.Range[rng.Columns[0].Rows[h + 1].ToString()].Value = "Ora " + h;

                if (informazioni.Count == 4)
                {
                    for (int j = 0; j < 4; j++)
                        _ws.Range[rng.Rows[1].Columns[j + 1].ToString()].Value = 15 * j + "-" + 15 * (j+1);
                }
                else
                    _ws.Range[rng.Cells[1,1].ToString()].Value = "0-60";

                if (_mercato != "MSD1")
                {
                    string mercatoPrec = Simboli.GetMercatoPrec(_mercato);

                    Excel.FormatCondition condGreater = _ws.Range[rng.Rows[2, rng.Rows.Count - 1].Columns[1, 4].ToString()].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: "=" + rng.Cells[2, 1] + " > '" + mercatoPrec + "'!" + rng.Cells[2, 1]);
                    condGreater.Interior.ColorIndex = Struct.COLORE_VARIAZIONE_POSITIVA;

                    Excel.FormatCondition condLess = _ws.Range[rng.Rows[2, rng.Rows.Count - 1].Columns[1, 4].ToString()].FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: "=" + rng.Cells[2, 1] + " < '" + mercatoPrec + "'!" + rng.Cells[2, 1]);
                    condLess.Interior.ColorIndex = Struct.COLORE_VARIAZIONE_NEGATIVA;
                }
                

            }
        }

        public override void UpdateData(bool all = true)
        {
            SplashScreen.UpdateStatus("Aggiorno informazioni");
            if (all)
            {
                CancellaDati();
                AggiornaDateTitoli();
            }
            CaricaInformazioni(all);            
        }
        private void CancellaDati()
        {
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "Gerarchia = '' OR Gerarchia IS NULL";

            foreach (DataRowView entita in categoriaEntita)
            {
                List<DataRow> entitaRif =
                   (from r in categoriaEntita.Table.AsEnumerable()
                    where r["Gerarchia"].Equals(entita["SiglaEntita"])
                    select r).ToList();

                int numEntita = Math.Max(entitaRif.Count, 1);

                for (int i = 0; i < numEntita; i++)
                {
                    Range rng = new Range(_definedNames.GetRowByName(entita["SiglaEntita"], "UM", "T") + 2, _definedNames.GetColFromName("RIF" + (i + 1), "PROGRAMMAQ1"), Date.GetOreGiorno(DataBase.DataAttiva) + 2, 4);

                    _ws.Range[rng.ToString()].Value = null;
                }
            }
        }
        public override void AggiornaDateTitoli()
        {
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "Gerarchia = '' OR Gerarchia IS NULL";

            foreach (DataRowView entita in categoriaEntita)
            {
                Range rng = new Range(_definedNames.GetRowByName(entita["SiglaEntita"], "DATA"), _struttura.colBlock + 1);
                _ws.Range[rng.ToString()].Value = DataBase.DataAttiva;
            }
        }
        public override void AggiornaGrafici()
        {
        }
        protected override void InsertPersonalizzazioni(object siglaEntita)
        {            
        }

        public override void CaricaInformazioni(bool all)
        {
            try
            {
                if (DataBase.OpenConnection())
                {
                    SplashScreen.UpdateStatus("Carico informazioni dal DB per " + _mercato);
                    _dataInizio = DataBase.DB.DataAttiva;

                    DataView datiApplicazioneH = DataBase.Select(DataBase.SP.APPLICAZIONE_INFORMAZIONE_H_EXPORT, "@IdApplicazione=" + Simboli.GetAppIDByMercato(_ws.Name) + ";@SiglaEntita=ALL;@SiglaCategoria=ALL;@DateFrom=" + _dataInizio.ToString("yyyyMMdd") + ";@DateTo=" + _dataInizio.ToString("yyyyMMdd")).DefaultView;

                    var listaEntitaInfo =
                        (from DataRowView r in datiApplicazioneH
                         group r by new { SiglaEntita = r["SiglaEntita"], SiglaInformazione = r["SiglaInformazione"], Riferimento = r["Riferimento"] } into g
                         select new { SiglaEntita = g.Key.SiglaEntita.ToString(), SiglaInformazione = g.Key.SiglaInformazione.ToString(), Riferimento = g.Key.Riferimento.ToString() }).ToList();

                    foreach (var entitaInfo in listaEntitaInfo)
                    {
                        SplashScreen.UpdateStatus("Scrivo informazioni " + entitaInfo.SiglaEntita);
                        datiApplicazioneH.RowFilter = "SiglaEntita = '" + entitaInfo.SiglaEntita + "' AND SiglaInformazione = '" + entitaInfo.SiglaInformazione + "' AND Riferimento = " + entitaInfo.Riferimento;

                        string quarter = Regex.Match(entitaInfo.SiglaInformazione, @"Q\d").Value;
                        quarter = quarter == "" ? "Q1" : quarter;

                        Range rng = new Range(_definedNames.GetRowByName(entitaInfo.SiglaEntita, "UM", "T") + 2, _definedNames.GetColFromName("RIF" + datiApplicazioneH[0]["Riferimento"], "PROGRAMMA" + quarter)).Extend(rowOffset: datiApplicazioneH.Count);

                        for (int i = 0; i <rng.Rows.Count; i++)
                            _ws.Range[rng.Rows[i].ToString()].Value = datiApplicazioneH[i]["Valore"];
                    }
                }
            }
            catch (Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "CaricaInformazioni InvioProgrammi SheetExport [all = " + all + "]: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public DateTime DataCaricamentoStruttura
        {
            get 
            {
                if (_dataCaricaStruttura.ContainsKey(_ws.Name))
                    return _dataCaricaStruttura[_ws.Name];
                else
                    return DateTime.MinValue;
            }
        }
    }
}
