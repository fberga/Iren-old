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
        public abstract void AggiornaRiepilogo(object entita, object azione, bool presente, DateTime? dataRif = null);
        public abstract void UpdateRiepilogo();
        public abstract void RiepilogoInEmergenza();

        #endregion
    }

    public class Riepilogo : ARiepilogo
    {
        #region Variabili

        protected Excel.Worksheet _ws;
        protected DefinedNames _nomiDefiniti;
        protected NewDefinedNames _newNomiDefiniti;
        protected int _rigaAttiva;
        protected int _colonnaInizio;
        protected int _nAzioni;
        protected static bool _resizeFatto = false;

        #endregion

        #region Costruttori

        public Riepilogo() : this((Excel.Worksheet)Utility.Workbook.WB.Sheets["Main"])  { }

        public Riepilogo(Excel.Worksheet ws)
        {
            _ws = ws;

            //dimensionamento celle in base ai parametri del DB
            DataView paramApplicazione = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE].DefaultView;

            _struttura = new Struct();
            _struttura.rigaBlock = 5;
            _struttura.colBlock = 59;
            _nomiDefiniti = new DefinedNames(_ws.Name);
            try
            {
                _newNomiDefiniti = new NewDefinedNames(_ws.Name);
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

            DataView azioni = DataBase.LocalDB.Tables[DataBase.Tab.AZIONE].DefaultView;
            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            DataView entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;

            categorie.RowFilter = "Operativa = 1";
            azioni.RowFilter = "Visibile = 1 AND Operativa = 1";
            entita.RowFilter = "";

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

        public override void InitLabels()
        {
            //inizializzo i label
            _ws.Shapes.Item("sfondo").LockAspectRatio = Office.MsoTriState.msoFalse;
            _ws.Shapes.Item("sfondo").Height = (float)(16.5 * _ws.Rows[5].Height);
            _ws.Shapes.Item("sfondo").LockAspectRatio = Office.MsoTriState.msoCTrue;

            _ws.Shapes.Item("lbTitolo").TextFrame.Characters().Text = Simboli.nomeApplicazione;
            _ws.Shapes.Item("lbDataInizio").TextFrame.Characters().Text = DataBase.DataAttiva.ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbDataFine").TextFrame.Characters().Text = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbVersione").TextFrame.Characters().Text = "Foglio v." + Utilities.WorkbookVersion.ToString();
            _ws.Shapes.Item("lbUtente").TextFrame.Characters().Text = "Utente: " + DataBase.LocalDB.Tables[DataBase.Tab.UTENTE].Rows[0]["Nome"];

            //aggiorna la scritta di modifica dati
            Simboli.ModificaDati = false;

            //aggiorna la scritta e il colore del label che mostra l'ambiente
            Simboli.Ambiente = ConfigurationManager.AppSettings["DB"];

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
            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            DataView entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            DataView azioni = DataBase.LocalDB.Tables[DataBase.Tab.AZIONE].DefaultView;

            //inserisco tutte le righe
            _newNomiDefiniti.AddName(_rigaAttiva++, "DATA");
            _newNomiDefiniti.AddName(_rigaAttiva++, "AZIONI_PADRE");
            _newNomiDefiniti.AddName(_rigaAttiva++, "AZIONI");

            foreach (DataRowView categoria in categorie)
            {
                _newNomiDefiniti.AddName(_rigaAttiva++, categoria["SiglaCategoria"]);
                entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "'";
                foreach (DataRowView e in entita)
                {
                    _newNomiDefiniti.AddName(_rigaAttiva, e["SiglaEntita"]);
                    _newNomiDefiniti.AddGOTO(e["SiglaEntita"], _rigaAttiva++, _colonnaInizio);
                }
            }
            
            //inserisco tutte le colonne
            _newNomiDefiniti.AddDate(_colonnaInizio++, "COLONNA_ENTITA");
            CicloGiorni((oreGiorno, suffissoData, giorno) => 
            {
                foreach (DataRowView azione in azioni)
                {
                    if (azione["Gerarchia"] != DBNull.Value)
                        _newNomiDefiniti.AddDate(_colonnaInizio++, azione["SiglaAzione"], suffissoData);
                }
            });
        }
        protected void InitBarraTitolo()
        {
            DataView entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            DataView azioni = DataBase.LocalDB.Tables[DataBase.Tab.AZIONE].DefaultView;

            Range rngTitleBar = new Range(_newNomiDefiniti.GetFirstRow(), _newNomiDefiniti.GetFirstCol() + 1, 3, azioni.Count);
            Range rngData = rngTitleBar.Cells[0, 0];
            Range rngAzioniPadre = rngTitleBar.Cells[1, 0];
            Range rngAzioni = rngTitleBar.Cells[2, 0];

            string azionePadre = "";
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                rngTitleBar.StartColumn = rngAzioni.StartColumn;
                _ws.Range[rngTitleBar.ToString()].Style = "recapTitleBarStyle";
                _ws.Range[rngTitleBar.ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

                foreach (DataRowView azione in azioni)
                {
                    if (!azione["Gerarchia"].Equals(azionePadre))
                    {
                        rngAzioniPadre.ColOffset = rngAzioni.StartColumn - rngAzioniPadre.StartColumn;
                        Style.RangeStyle(_ws.Range[rngAzioniPadre.ToString()], "Merge:true;FontSize:9");
                        _ws.Range[rngAzioniPadre.ToString()].Value = azionePadre;
                        azionePadre = azione["Gerarchia"].ToString();
                        rngAzioniPadre.StartColumn = rngAzioni.StartColumn;
                    }
                    _ws.Range[rngAzioni.ToString()].Value = azione["DesAzioneBreve"];
                    Style.RangeStyle(_ws.Range[rngAzioni.ToString()], "FontSize:7");
                    rngAzioni.StartColumn++;
                }
                rngAzioniPadre.ColOffset = rngAzioni.StartColumn - rngAzioniPadre.StartColumn;
                Style.RangeStyle(_ws.Range[rngAzioniPadre.ToString()], "Merge:true;FontSize:9");
                _ws.Range[rngAzioniPadre.ToString()].Value = azionePadre;
                azionePadre = "";
                rngAzioniPadre.StartColumn = rngAzioni.StartColumn;

                rngData.ColOffset = rngAzioni.StartColumn - rngData.StartColumn;
                Style.RangeStyle(_ws.Range[rngData.ToString()], "Merge:true;FontSize:10;NumberFormat:[ddd d mmm yyyy]");
                _ws.Range[rngData.ToString()].Value = giorno;
                rngData.StartColumn = rngAzioni.StartColumn;
            });
        }
        protected void FormattaAllDati()
        {
            DataView azioni = DataBase.LocalDB.Tables[DataBase.Tab.AZIONE].DefaultView;
            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            DataView entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;

            Range rngAll = new Range(_newNomiDefiniti.GetFirstRow(), _newNomiDefiniti.GetFirstCol() + 1, _newNomiDefiniti.GetRowOffset(), _newNomiDefiniti.GetColOffsetRiepilogo() - 1);
            Range rngData = new Range(_newNomiDefiniti.GetFirstRow() + 3, _newNomiDefiniti.GetFirstCol(), _newNomiDefiniti.GetRowOffset() - 3, _newNomiDefiniti.GetColOffsetRiepilogo());
            
            _ws.Range[rngData.ToString()].Style = "recapAllDatiStyle";
            _ws.Range[rngData.Columns[0].ToString()].Style = "recapEntityBarStyle";
            _ws.Range[rngData.Columns[0].ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

            Excel.Range xlrng = _ws.Range[rngAll.Rows[1, rngAll.Rows.Count].ToString()];
            //trovo tutte le aree unite e creo il blocco col bordo grosso
            int i = 0;
            int colspan = 0;
            while (i < xlrng.Columns.Count)
            {
                colspan = xlrng.Cells[1, i + 1].MergeArea().Columns.Count;
                _ws.Range[rngAll.Columns[i, i + colspan].ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
                _ws.Range[rngAll.Columns[i, i + colspan].ToString()].Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                i += colspan;
            }
        }
        protected void InitBarraEntita()
        {
            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            DataView entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;

            foreach (DataRowView categoria in categorie)
            {
                Range rng = new Range(_newNomiDefiniti.GetRowByName(categoria["SiglaCategoria"]), _newNomiDefiniti.GetFirstCol(), 1, _newNomiDefiniti.GetColOffsetRiepilogo());
                Style.RangeStyle(_ws.Range[rng.ToString()], "Style:recapCategoryTitle;Borders:[left:medium,top:medium,right:medium]");
                _ws.Range[rng.Columns[0].ToString()].Value = categoria["DesCategoria"];
                entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "'";
                foreach (DataRowView e in entita)
                {
                    rng.StartRow++;
                    _ws.Range[rng.Columns[0].ToString()].Value = (e["Gerarchia"] is DBNull ? "" : "     ") + e["DesEntita"];
                    _ws.Range[rng.Columns[0].ToString()].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;
                }
            }
            
            _ws.Columns[_struttura.colRecap].EntireColumn.AutoFit();

            //categorie.RowFilter = "Operativa = 1";
            //entita.RowFilter = "";

            //object[,] values = new object[categorie.Count + entita.Count, 1];
            //Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, _colonnaInizio], _ws.Cells[_rigaAttiva + values.Length - 1, _colonnaInizio]];
            //rng.Style = "recapEntityBarStyle";
            //Style.RangeStyle(rng, "borders:[top:medium,right:medium,bottom:medium,left:medium,insideh:thin]");
            //int i = 0;
            //foreach (DataRowView categoria in categorie)
            //{
            //    Excel.Range titoloCategoria = _ws.Range[_ws.Cells[_rigaAttiva + i, _colonnaInizio], _ws.Cells[_rigaAttiva + i, _colonnaInizio + ((Struct.intervalloGiorni + 1) * _nAzioni)]];
            //    titoloCategoria.Merge();
            //    titoloCategoria.Style = "recapCategoryTitle";
            //    Style.RangeStyle(titoloCategoria, "Borders:[left:medium, top:medium, right:medium]");

            //    values[i++, 0] = categoria["DesCategoria"];
            //    entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "'";

                

            //    foreach (DataRowView e in entita)
            //    {
            //        _nomiDefiniti.Add(DefinedNames.GetName("RIEPILOGO", e["siglaEntita"], "GOTO"), Tuple.Create(_rigaAttiva + i, _colonnaInizio));
            //        values[i++, 0] = (e["Gerarchia"] is DBNull ? "" : "     ") + e["DesEntita"];
            //    }
            //}
            //rng.Value = values;

            //categorie.RowFilter = "";
            //entita.RowFilter = "";
            //rng.EntireColumn.AutoFit();
        }
        protected void AbilitaAzioni()
        {
            DataView entitaAzioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAAZIONE].DefaultView;
            entitaAzioni.RowFilter = "";

            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                foreach (DataRowView azione in entitaAzioni)
                {                    
                    Range cellaAzione = new Range(_newNomiDefiniti.GetRowByName(azione["SiglaEntita"]), _newNomiDefiniti.GetColFromName(azione["SiglaAzione"], suffissoData));
                    _ws.Range[cellaAzione.ToString()].Interior.Pattern = Excel.XlPattern.xlPatternNone;
                }
            });
        }
        protected void CaricaDatiRiepilogo()
        {
            try
            {
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    DataView datiRiepilogo = DataBase.DB.Select(DataBase.SP.APPLICAZIONE_RIEPILOGO, "@Data=" + giorno.ToString("yyyyMMdd")).DefaultView;
                    foreach (DataRowView valore in datiRiepilogo)
                    {
                        Range cellaAzione = new Range(_newNomiDefiniti.GetRowByName(valore["SiglaEntita"]), _newNomiDefiniti.GetColFromName(valore["SiglaAzione"], suffissoData));                        
                        string commento = "";

                        Excel.Range rng = _ws.Range[cellaAzione.ToString()];

                        if (valore["Presente"].Equals("1"))
                        {
                            rng.ClearComments();
                            DateTime data = DateTime.ParseExact(valore["Data"].ToString(), "yyyyMMddHHmm", CultureInfo.InvariantCulture);
                            commento = "Utente: " + valore["Utente"] + "\nData: " + data.ToString("dd MMM yyyy") + "\nOra: " + data.ToString("HH:mm");
                            rng.AddComment(commento);
                            rng.Value = "OK";
                            Style.RangeStyle(rng, "BackColor:4;Align:Center");
                        }
                        else
                        {
                            rng.ClearComments();
                            rng.Value = "Non presente";
                            Style.RangeStyle(rng, "ForeColor:3;Bold:False;Align:Center;FontSize:7");
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

        public override void AggiornaRiepilogo(object entita, object azione, bool presente, DateTime? dataRif = null)
        {
            if(dataRif == null)
                dataRif = DataBase.DB.DataAttiva;

            if (_nomiDefiniti.IsDefined(DefinedNames.GetName("RIEPILOGO", entita, azione, Date.GetSuffissoData(dataRif.Value))))
            {
                Tuple<int, int> cella = _nomiDefiniti["RIEPILOGO", entita, azione, Date.GetSuffissoData(dataRif.Value)][0];
                Excel.Range rng = _ws.Cells[cella.Item1, cella.Item2];

                if (presente)
                {
                    string commento = "Utente: " + DataBase.LocalDB.Tables[DataBase.Tab.UTENTE].Rows[0]["Nome"] + "\nData: " + DateTime.Now.ToString("dd MMM yyyy") + "\nOra: " + DateTime.Now.ToString("HH:mm");
                    rng.ClearComments();
                    rng.AddComment(commento).Visible = false;
                    rng.Value = "OK";
                    Style.RangeStyle(rng, "FontSize:9;ForeColor:1;BackColor:4;Align:Center;Bold:true");
                }
                else
                {
                    rng.Value = "Non presente";
                    Style.RangeStyle(rng, "FontSize:7;ForeColor:3;BackColor:2;Align:Center;Bold:false");
                }
            }
        }

        protected void AggiornaDate()
        {
            _ws.Shapes.Item("lbDataInizio").TextFrame.Characters().Text = DataBase.DB.DataAttiva.ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbDataFine").TextFrame.Characters().Text = DataBase.DB.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("ddd d MMM yyyy");

            CicloGiorni((oreGiorno, suffissoData, giorno) => 
            {
                if (_nomiDefiniti.IsDefined(DefinedNames.GetName("RIEPILOGO", "T", suffissoData)))
                {
                    Tuple<int, int>[] riga = _nomiDefiniti.GetRanges(DefinedNames.GetName("RIEPILOGO", "T", suffissoData))[0];
                    //TODO Sistemare
                    //_ws.Range[_nomiDefiniti.GetRange(riga)].Value = giorno;
                }
            });
        }
        public override void UpdateRiepilogo()
        {
            AggiornaDate();
            AbilitaAzioni();
            CaricaDatiRiepilogo();
        }

        protected void DisabilitaTutto()
        {
            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            DataView entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;

            categorie.RowFilter = "Operativa = 1";

            foreach (DataRowView categoria in categorie)
            {
                entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "'";
                
                foreach (DataRowView e in entita)
                {
                    if (_nomiDefiniti.IsDefined(DefinedNames.GetName("RIEPILOGO", e["siglaEntita"])))
                    {
                        Tuple<int, int>[] riga = _nomiDefiniti["RIEPILOGO", e["siglaEntita"], Simboli.EXCLUDE, "GOTO"];

                        Excel.Range rng = _ws.Range[_ws.Cells[riga[0].Item1, riga[0].Item2], _ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]];
                        rng.Value = "";
                        rng.ClearComments();

                        Style.RangeStyle(rng, "BackPattern: CrissCross");
                    }
                }
            }

        }
        public override void RiepilogoInEmergenza()
        {
            AggiornaDate();
            DisabilitaTutto();
        }

        #endregion

    }
}
