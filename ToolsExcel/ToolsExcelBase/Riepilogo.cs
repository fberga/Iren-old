using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Data;
using System.Globalization;
using System.Configuration;
using System.Windows.Forms;
using Iren.ToolsExcel.Utility;

namespace Iren.ToolsExcel.Base
{
    public abstract class ARiepilogo
    {
        #region Variabili

        protected Struct _struttura;
        protected DataView _azioni = new DataView(DataBase.LocalDB.Tables[DataBase.Tab.AZIONE]);
        protected DataView _categorie = new DataView(DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA]);
        protected DataView _entita = new DataView(DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA]);
        protected DataView _entitaAzioni = new DataView(DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE]);

        #endregion

        #region Metodi

        protected void CicloGiorni(Action<int, string, DateTime> callback)
        {
            DateTime dataInizio = DataBase.DataAttiva;
            DateTime dataFine = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni);
            CicloGiorni(dataInizio, dataFine, callback);
        }
        protected void CicloGiorni(DateTime dataInizio, DateTime dataFine, Action<int, string, DateTime> callback)
        {
            for (DateTime giorno = dataInizio; giorno <= dataFine; giorno = giorno.AddDays(1))
            {
                int oreGiorno = Date.GetOreGiorno(giorno);
                string suffissoData = Date.GetSuffissoData(dataInizio, giorno);

                if (giorno == dataInizio && _struttura.visData0H24)
                {
                    oreGiorno++;
                }

                callback(oreGiorno, suffissoData, giorno);
            }
        }

        public abstract void InitLabels();
        public abstract void LoadStructure();
        public void AggiornaRiepilogo(object siglaEntita, object siglaAzione, bool presente)
        {
            AggiornaRiepilogo(siglaEntita, siglaAzione, presente, DataBase.DataAttiva);
        }
        public abstract void AggiornaRiepilogo(object siglaEntita, object siglaAzione, bool presente, DateTime dataRif);
        public abstract void UpdateData();
        public abstract void RiepilogoInEmergenza();

        #endregion
    }

    public class Riepilogo : ARiepilogo
    {
        #region Variabili

        protected Excel.Worksheet _ws;
        protected DefinedNames _definedNames;
        protected int _rigaAttiva;
        protected int _colonnaInizio;
        protected int _nAzioni;
        protected static bool _resizeFatto = false;

        #endregion

        #region Costruttori

        public Riepilogo() : this(Utility.Workbook.Main)  { }

        public Riepilogo(Excel.Worksheet ws)
        {
            _ws = ws;

            //dimensionamento celle in base ai parametri del DB
            DataView paramApplicazione = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE].DefaultView;

            _struttura = new Struct();
            _struttura.rigaBlock = 5;
            _struttura.colBlock = 59;
            try
            {
                _definedNames = new DefinedNames(_ws.Name);
            }
            catch
            {

            }
        }

        #endregion

        #region Metodi

        public override void LoadStructure()
        {
            _colonnaInizio = _struttura.colRecap;
            _rigaAttiva = _struttura.rowRecap;

            InitLabels();
            Clear();

            if (Struct.visualizzaRiepilogo)
            {
                _categorie.RowFilter = "Operativa = 1";
                _azioni.RowFilter = "Visibile = 1 AND Operativa = 1";
                _entita.RowFilter = "";

                CreaNomiCelle();
                InitBarraTitolo();
                _rigaAttiva += 3;
                FormattaAllDati();
                InitBarraEntita();
                AbilitaAzioni();
                CaricaDatiRiepilogo();

                //Se sono in multiscreen lascio il riepilogo alla fine, altrimenti lo riporto all'inizio
                if (Screen.AllScreens.Length == 1)
                {
                    _ws.Application.ActiveWindow.SmallScroll(Type.Missing, Type.Missing, _struttura.colRecap - _struttura.colBlock - 1);
                }
            }
        }

        public override void InitLabels()
        {
            //inizializzo dimensione e posizione di default dei label
            //_ws.Shapes.Item("sfondo").LockAspectRatio = Office.MsoTriState.msoFalse;
            //_ws.Shapes.Item("sfondo").Height = (float)(16.5 * _ws.Rows[5].Height);
            //_ws.Shapes.Item("sfondo").LockAspectRatio = Office.MsoTriState.msoCTrue;

            _ws.Shapes.Item("lbTitolo").TextFrame.Characters().Text = Simboli.nomeApplicazione;
            _ws.Shapes.Item("lbDataInizio").TextFrame.Characters().Text = DataBase.DataAttiva.ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbDataFine").TextFrame.Characters().Text = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbVersione").TextFrame.Characters().Text = "Foglio v." + Utility.Workbook.WorkbookVersion.ToString();
            _ws.Shapes.Item("lbUtente").TextFrame.Characters().Text = "Utente: " + DataBase.LocalDB.Tables[DataBase.Tab.UTENTE].Rows[0]["Nome"];

            DataView applicazione = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE].DefaultView;
           
            //applico colori da DB
            _ws.Shapes.Item("lbTitolo").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbLinee[0], Simboli.rgbLinee[1], Simboli.rgbLinee[2]));
            _ws.Shapes.Item("lbTitolo").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbTitolo[0], Simboli.rgbTitolo[1], Simboli.rgbTitolo[2]));
            _ws.Shapes.Item("sfondo").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbLinee[0], Simboli.rgbLinee[1], Simboli.rgbLinee[2]));
            _ws.Shapes.Item("sfondo").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbSfondo[0], Simboli.rgbSfondo[1], Simboli.rgbSfondo[2]));
            _ws.Shapes.Item("lbDataInizio").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbTitolo[0], Simboli.rgbTitolo[1], Simboli.rgbTitolo[2]));
            _ws.Shapes.Item("lbDataFine").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbTitolo[0], Simboli.rgbTitolo[1], Simboli.rgbTitolo[2]));
            _ws.Shapes.Item("lbMercato").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(Simboli.rgbTitolo[0], Simboli.rgbTitolo[1], Simboli.rgbTitolo[2]));

            //aggiorna la scritta di modifica dati
            Simboli.ModificaDati = false;

            //aggiorna la scritta e il colore del label che mostra l'ambiente
            Simboli.Ambiente = Utility.Workbook.AppSettings("DB");

            if (Struct.intervalloGiorni > 0)
            {
                _ws.Shapes.Item("lbDataInizio").LockAspectRatio = Office.MsoTriState.msoFalse;
                _ws.Shapes.Item("lbDataInizio").Width = 26 * (float)_ws.Columns[1].Width;
                _ws.Shapes.Item("lbDataFine").Visible = Office.MsoTriState.msoTrue;
                _ws.Shapes.Item("lbDataInizio").LockAspectRatio = Office.MsoTriState.msoTrue;
            }
            else
            {
                _ws.Shapes.Item("lbDataInizio").LockAspectRatio = Office.MsoTriState.msoFalse;
                _ws.Shapes.Item("lbDataInizio").Width = 54 * (float)_ws.Columns[1].Width;
                _ws.Shapes.Item("lbDataFine").Visible = Office.MsoTriState.msoFalse;
                _ws.Shapes.Item("lbDataInizio").LockAspectRatio = Office.MsoTriState.msoTrue;
            }
        }
        protected virtual void Clear()
        {
            _ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            _ws.UsedRange.EntireColumn.Delete();
            _ws.UsedRange.FormatConditions.Delete();
            _ws.UsedRange.Interior.ColorIndex = 2;
            _ws.UsedRange.EntireRow.Hidden = false;
            _ws.UsedRange.Font.Size = 8;
            _ws.UsedRange.NumberFormat = "General";
            _ws.UsedRange.Font.Name = "Verdana";

            _ws.Columns.ColumnWidth = 9;

            _ws.Range[Range.GetRange(1, 1, 1, _struttura.colRecap - 1)].EntireColumn.ColumnWidth = Struct.cell.width.empty;            
            _ws.Rows[1].RowHeight = Struct.cell.height.empty;

            ((Excel._Worksheet)_ws).Activate();
            _ws.Application.ActiveWindow.FreezePanes = false;
            _ws.Cells[_struttura.rigaBlock, _struttura.colBlock].Select();
            _ws.Application.ActiveWindow.ScrollColumn = 1;
            _ws.Application.ActiveWindow.ScrollRow = 1;
            _ws.Application.ActiveWindow.FreezePanes = true;
        }
        protected virtual void CreaNomiCelle()
        {
            //inserisco tutte le righe
            _definedNames.AddName(_rigaAttiva++, "DATA");
            _definedNames.AddName(_rigaAttiva++, "AZIONI_PADRE");
            _definedNames.AddName(_rigaAttiva++, "AZIONI");

            foreach (DataRowView categoria in _categorie)
            {
                _definedNames.AddName(_rigaAttiva++, categoria["SiglaCategoria"]);
                _entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "'";
                foreach (DataRowView e in _entita)
                {
                    _definedNames.AddName(_rigaAttiva, e["SiglaEntita"]);
                    _definedNames.AddGOTO(e["SiglaEntita"], Range.R1C1toA1(_rigaAttiva++, _colonnaInizio));
                }
            }
            
            //inserisco tutte le colonne
            _definedNames.AddCol(_colonnaInizio++, "COLONNA_ENTITA");
            CicloGiorni((oreGiorno, suffissoData, giorno) => 
            {
                foreach (DataRowView azione in _azioni)
                {
                    if (azione["Gerarchia"] != DBNull.Value)
                        _definedNames.AddCol(_colonnaInizio++, azione["SiglaAzione"], suffissoData);
                }
            });
            _definedNames.DumpToDataSet();
        }
        protected void InitBarraTitolo()
        {
            Range rngTitleBar = new Range(_definedNames.GetFirstRow(), _definedNames.GetFirstCol() + 1, 3, _azioni.Count);
            Range rngData = rngTitleBar.Cells[0, 0];
            Range rngAzioniPadre = rngTitleBar.Cells[1, 0];
            Range rngAzioni = rngTitleBar.Cells[2, 0];

            string azionePadre = "";
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                rngTitleBar.StartColumn = rngAzioni.StartColumn;
                _ws.Range[rngTitleBar.ToString()].Style = "recapTitleBarStyle";
                _ws.Range[rngTitleBar.ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

                foreach (DataRowView azione in _azioni)
                {
                    if (!azione["Gerarchia"].Equals(azionePadre))
                    {
                        rngAzioniPadre.ColOffset = rngAzioni.StartColumn - rngAzioniPadre.StartColumn;
                        Style.RangeStyle(_ws.Range[rngAzioniPadre.ToString()], merge:true, fontSize:9);
                        _ws.Range[rngAzioniPadre.ToString()].Value = azionePadre;
                        azionePadre = azione["Gerarchia"].ToString();
                        rngAzioniPadre.StartColumn = rngAzioni.StartColumn;
                    }
                    _ws.Range[rngAzioni.ToString()].Value = azione["DesAzioneBreve"];
                    Style.RangeStyle(_ws.Range[rngAzioni.ToString()], fontSize:7);
                    rngAzioni.StartColumn++;
                }
                rngAzioniPadre.ColOffset = rngAzioni.StartColumn - rngAzioniPadre.StartColumn;
                Style.RangeStyle(_ws.Range[rngAzioniPadre.ToString()], merge:true, fontSize:9);
                _ws.Range[rngAzioniPadre.ToString()].Value = azionePadre;
                azionePadre = "";
                rngAzioniPadre.StartColumn = rngAzioni.StartColumn;

                rngData.ColOffset = rngAzioni.StartColumn - rngData.StartColumn;
                Style.RangeStyle(_ws.Range[rngData.ToString()], merge:true, fontSize:10, numberFormat:"ddd d mmm yyyy");
                _ws.Range[rngData.ToString()].Value = giorno;
                rngData.StartColumn = rngAzioni.StartColumn;
            });
        }
        protected void FormattaAllDati()
        {
            Range rngAll = new Range(_definedNames.GetFirstRow(), _definedNames.GetFirstCol() + 1, _definedNames.GetRowOffset(), _definedNames.GetColOffsetRiepilogo() - 1);
            Range rngData = new Range(_definedNames.GetFirstRow() + 3, _definedNames.GetFirstCol(), _definedNames.GetRowOffset() - 3, _definedNames.GetColOffsetRiepilogo());
            
            _ws.Range[rngData.ToString()].Style = "recapAllDatiStyle";
            _ws.Range[rngData.Columns[0].ToString()].Style = "recapEntityBarStyle";
            _ws.Range[rngData.Columns[0].ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

            Excel.Range xlrng = _ws.Range[rngAll.Rows[1, rngAll.Rows.Count - 1].ToString()];
            //trovo tutte le aree unite e creo il blocco col bordo grosso
            int i = 0;
            int colspan = 0;
            while (i < xlrng.Columns.Count)
            {
                colspan = xlrng.Cells[1, i + 1].MergeArea().Columns.Count;
                _ws.Range[rngAll.Columns[i, colspan - 1].ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
                _ws.Range[rngAll.Columns[i, colspan - 1].ToString()].Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                i += colspan;
            }
        }
        protected void InitBarraEntita()
        {
            foreach (DataRowView categoria in _categorie)
            {
                Range rng = new Range(_definedNames.GetRowByName(categoria["SiglaCategoria"]), _definedNames.GetFirstCol(), 1, _definedNames.GetColOffsetRiepilogo());
                Style.RangeStyle(_ws.Range[rng.ToString()], style: "recapCategoryTitle", borders: "[left:medium,top:medium,right:medium]", merge: true);
                _ws.Range[rng.Columns[0].ToString()].Value = categoria["DesCategoria"];
                _entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "'";
                foreach (DataRowView entita in _entita)
                {
                    rng.StartRow++;
                    _ws.Range[rng.Columns[0].ToString()].Value = (entita["Gerarchia"] is DBNull ? "" : "     ") + entita["DesEntita"];
                    _ws.Range[rng.Columns[0].ToString()].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                }
            }
            
            _ws.Columns[_struttura.colRecap].EntireColumn.AutoFit();
        }
        protected void AbilitaAzioni()
        {
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                foreach (DataRowView azione in _entitaAzioni)
                {                    
                    Range cellaAzione = new Range(_definedNames.GetRowByName(azione["SiglaEntita"]), _definedNames.GetColFromName(azione["SiglaAzione"], suffissoData));
                    _ws.Range[cellaAzione.ToString()].Interior.Pattern = Excel.XlPattern.xlPatternNone;
                    _ws.Range[cellaAzione.ToString()].Interior.ColorIndex = 2;
                }
            });
        }
        protected void CaricaDatiRiepilogo()
        {
            try
            {
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    if (DataBase.OpenConnection())
                    {
                        DataView datiRiepilogo = DataBase.Select(DataBase.SP.APPLICAZIONE_RIEPILOGO, "@Data=" + giorno.ToString("yyyyMMdd")).DefaultView;
                        foreach (DataRowView valore in datiRiepilogo)
                        {
                            Range cellaAzione = new Range(_definedNames.GetRowByName(valore["SiglaEntita"]), _definedNames.GetColFromName(valore["SiglaAzione"], suffissoData));

                            Excel.Range rng = _ws.Range[cellaAzione.ToString()];

                            if (valore["Presente"].Equals("1"))
                            {
                                rng.ClearComments();
                                DateTime data = DateTime.ParseExact(valore["Data"].ToString(), "yyyyMMddHHmm", CultureInfo.InvariantCulture);
                                rng.AddComment("Utente: " + valore["Utente"] + "\nData: " + data.ToString("dd MMM yyyy") + "\nOra: " + data.ToString("HH:mm"));
                                rng.Value = "OK";
                                Style.RangeStyle(rng, foreColor: 1, bold: true, fontSize: 9, backColor: 4, align: Excel.XlHAlign.xlHAlignCenter);
                            }
                            else
                            {
                                rng.ClearComments();
                                rng.Value = "Non presente";
                                Style.RangeStyle(rng, foreColor: 3, bold: false, fontSize: 7, backColor: 2, align: Excel.XlHAlign.xlHAlignCenter);
                            }
                        }
                    }
                });
            }
            catch (Exception e)
            {
                Utility.Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "CaricaDatiRiepilogo: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public override void AggiornaRiepilogo(object siglaEntita, object siglaAzione, bool presente, DateTime dataRif)
        {

            if (dataRif - DataBase.DataAttiva <= new TimeSpan(Struct.intervalloGiorni, 0, 0, 0))
            {
                if (Struct.visualizzaRiepilogo && !Simboli.EmergenzaForzata)
                {
                    Range cell = _definedNames.Get(siglaEntita, siglaAzione, Date.GetSuffissoData(dataRif));
                    Excel.Range rng = _ws.Range[cell.ToString()];
                    if (presente)
                    {
                        string commento = "Utente: " + DataBase.LocalDB.Tables[DataBase.Tab.UTENTE].Rows[0]["Nome"] + "\nData: " + DateTime.Now.ToString("dd MMM yyyy") + "\nOra: " + DateTime.Now.ToString("HH:mm");
                        rng.ClearComments();
                        rng.AddComment(commento).Visible = false;
                        rng.Value = "OK";
                        Style.RangeStyle(rng, foreColor: 1, bold: true, fontSize: 9, backColor: 4, align: Excel.XlHAlign.xlHAlignCenter);
                    }
                    else
                    {
                        rng.ClearComments();
                        rng.Value = "Non presente";
                        Style.RangeStyle(rng, foreColor: 3, bold: false, fontSize: 7, backColor: 2, align: Excel.XlHAlign.xlHAlignCenter);
                    }
                }
            }
        }

        private void CancellaDati()
        {
            Range rngData = new Range(_definedNames.GetFirstRow() + 3, _definedNames.GetFirstCol() + 1, _definedNames.GetRowOffset() - 3, _definedNames.GetColOffsetRiepilogo() - 1);
            _ws.Range[rngData.ToString()].Value = null;
            _ws.Range[rngData.ToString()].Interior.ColorIndex = 2;
            _ws.Range[rngData.ToString()].ClearComments();
        }
        protected void AggiornaDate()
        {
            _ws.Shapes.Item("lbDataInizio").TextFrame.Characters().Text = DataBase.DB.DataAttiva.ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbDataFine").TextFrame.Characters().Text = DataBase.DB.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("ddd d MMM yyyy");
            
            if (Struct.visualizzaRiepilogo)
            {
                _azioni.RowFilter = "Visibile = 1 AND Operativa = 1 AND Gerarchia IS NOT NULL";
            
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    Range cell = new Range(_definedNames.GetRowByName("DATA"), _definedNames.GetColFromName(_azioni[0]["SiglaAzione"], suffissoData));
                    _ws.Range[cell.ToString()].Value = giorno;
                });
                _azioni.RowFilter = "Visibile = 1 AND Operativa = 1";
            }
        }
        public override void UpdateData()
        {
            if (_definedNames != null)
            {
                AggiornaDate();

                if (Struct.visualizzaRiepilogo)
                {
                    CancellaDati();
                    AbilitaAzioni();
                    CaricaDatiRiepilogo();
                }
            }
        }

        protected void DisabilitaTutto()
        {
            Range rngData = new Range(_definedNames.GetFirstRow() + 3, _definedNames.GetFirstCol() + 1, _definedNames.GetRowOffset() - 3, _definedNames.GetColOffsetRiepilogo() - 1);

            Style.RangeStyle(_ws.Range[rngData.ToString()], pattern: Excel.XlPattern.xlPatternCrissCross);
        }
        public override void RiepilogoInEmergenza()
        {
            if (Struct.visualizzaRiepilogo)
            {
                AggiornaDate();
                DisabilitaTutto();
            }
        }

        #endregion

    }
}
