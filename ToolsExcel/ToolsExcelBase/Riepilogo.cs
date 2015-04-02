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

            _ws.Range[_ws.Cells[1, 1], _ws.Cells[1, _struttura.colRecap - 1]].EntireColumn.ColumnWidth = Struct.cell.width.empty;            
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
                _newNomiDefiniti.AddName(_rigaAttiva++, categoria["SiglaCategoria"], "TITOLO");
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
                        _newNomiDefiniti.AddDate(_colonnaInizio++, azione["Gerarchia"], azione["SiglaAzione"], suffissoData);
                }
            });
        }
        protected void InitBarraTitolo()
        {
            DataView entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            DataView azioni = DataBase.LocalDB.Tables[DataBase.Tab.AZIONE].DefaultView;

            Range rngTitleBar = new Range(_newNomiDefiniti.GetRowByName("DATA"), _newNomiDefiniti.GetColFromName(azioni[0]["Gerarchia"], azioni[0]["SiglaAzione"], Date.GetSuffissoData(DataBase.DataAttiva)), 3, azioni.Count);
            Range rngData = new Range(rngTitleBar.StartRow, rngTitleBar.StartColumn);
            Range rngAzioniPadre = new Range(_newNomiDefiniti.GetRowByName("AZIONI_PADRE"), rngTitleBar.StartColumn);
            Range rngAzioni = new Range(_newNomiDefiniti.GetRowByName("AZIONI"), rngTitleBar.StartColumn);

            string azionePadre = "";
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                foreach (DataRowView azione in azioni)
                {
                    if (!azione["Gerarchia"].Equals(azionePadre))
                    {
                        rngAzioniPadre.ColOffset = rngAzioni.StartColumn - rngAzioniPadre.StartColumn;
                        _ws.Range[rngAzioniPadre.ToString()].Merge();
                        _ws.Range[rngAzioniPadre.ToString()].Value = azionePadre;
                        azionePadre = azione["Gerarchia"].ToString();
                        rngAzioniPadre.StartColumn = rngAzioni.StartColumn;
                    }
                    _ws.Range[rngAzioni.ToString()].Value = azione["DesAzioneBreve"];
                    rngAzioni.StartColumn++;
                }
                rngAzioniPadre.ColOffset = rngAzioni.StartColumn - rngAzioniPadre.StartColumn;
                _ws.Range[rngAzioniPadre.ToString()].Merge();
                _ws.Range[rngAzioniPadre.ToString()].Value = azionePadre;
                azionePadre = "";
                rngAzioniPadre.StartColumn = rngAzioni.StartColumn;

                rngData.ColOffset = rngAzioni.StartColumn - rngData.StartColumn;
                _ws.Range[rngData.ToString()].Merge();
                _ws.Range[rngData.ToString()].Value = giorno;
                rngData.StartColumn = rngAzioni.StartColumn;

                _ws.Range[rngTitleBar.ToString()].Style = "recapTitleBarStyle";
                Style.RangeStyle(_ws.Range[rngTitleBar.ToString()], "FontSize:10;NumberFormat:[ddd d mmm yyyy]");
                rngTitleBar.StartColumn = rngAzioni.StartColumn;
            });
            


            //_nAzioni = 0;
            //Dictionary<object, List<object>> valAzioni = new Dictionary<object, List<object>>();
            //foreach (DataRowView azione in azioni)
            //{
            //    if (!valAzioni.ContainsKey(azione["Gerarchia"]))
            //    {
            //        valAzioni.Add(azione["Gerarchia"], new List<object>() { azione["DesAzioneBreve"] });
            //    }
            //    else
            //    {
            //        valAzioni[azione["Gerarchia"]].Add(azione["DesAzioneBreve"]);
            //    }
            //    _nAzioni++;
            //}            
            //int nAzioniPadre = valAzioni.Count;

            ////numero totale di celle della barra del titolo
            //object[,] values = new object[3, _nAzioni];
            ////la prima libera per mettere la data successivamente
            //int[] azioniPerPadre = new int[valAzioni.Count];
            //int ipadre = 0;
            //int iazioni = 0;
            //int j = 0;
            //foreach (KeyValuePair<object, List<object>> keyVal in valAzioni)
            //{
            //    azioniPerPadre[j++] = keyVal.Value.Count;
            //    values[1, ipadre] = keyVal.Key.ToString().ToUpperInvariant();
            //    ipadre += keyVal.Value.Count;
            //    foreach (object nomeAzione in keyVal.Value) 
            //        values[2, iazioni++] = nomeAzione.ToString().ToUpperInvariant();
            //}

            //int colonnaInizio = _colonnaInizio + 1;
            //CicloGiorni((oreGiorno, suffissoData, giorno) =>
            //{
            //    values[0, 0] = giorno;
            //    Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, colonnaInizio], _ws.Cells[_rigaAttiva + 2, colonnaInizio + _nAzioni - 1]];

            //    _nomiDefiniti.Add(DefinedNames.GetName("RIEPILOGO", "T", suffissoData), _rigaAttiva, colonnaInizio, _rigaAttiva, colonnaInizio + _nAzioni - 1);

            //    rng.Style = "recapTitleBarStyle";
            //    rng.Rows[1].Merge();
            //    Style.RangeStyle(rng.Rows[1], "FontSize:10;NumberFormat:[ddd d mmm yyyy]");
            //    int i = 1;
            //    foreach (int numAzioniPerPadre in azioniPerPadre)
            //    {
            //        _ws.Range[rng.Cells[2, i], rng.Cells[2, i + numAzioniPerPadre - 1]].Merge();
            //        Style.RangeStyle(_ws.Range[rng.Cells[3, i], rng.Cells[3, i + numAzioniPerPadre - 1]], "FontSize:7;Borders:[left:medium, bottom:medium, right:medium, insidev:thin]");
            //        i += numAzioniPerPadre;
            //    }
            //    rng.Value = values;

            //    colonnaInizio += _nAzioni;
            //});
        }
        protected void FormattaAllDati()
        {
            DataView azioni = DataBase.LocalDB.Tables[DataBase.Tab.AZIONE].DefaultView;
            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            DataView entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;

            azioni.RowFilter = "Gerarchia IS NOT NULL";
            categorie.RowFilter = "Operativa = 1";
            entita.RowFilter = "";

            int numRighe = categorie.Count + entita.Count - 1;
            int colonnaInizio = _colonnaInizio + 1;

            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, colonnaInizio], _ws.Cells[_rigaAttiva + numRighe, colonnaInizio + _nAzioni - 1]];
                rng.Style = "recapAllDatiStyle";
                rng.BorderAround2(Type.Missing, Excel.XlBorderWeight.xlMedium);
                rng.ColumnWidth = 9;

                int j = 0;
                string gerarchia = azioni[0]["Gerarchia"].ToString();
                foreach(DataRowView azione in azioni) 
                {
                    if (!gerarchia.Equals(azione["Gerarchia"]))
                    {
                        gerarchia = azione["Gerarchia"].ToString();
                        rng.Columns[j].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                    }

                    int i = 0;
                    foreach (DataRowView categoria in categorie)
                    {
                        i++;
                        entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "'";
                        foreach (DataRowView e in entita)
                        {
                            _nomiDefiniti.Add(DefinedNames.GetName("RIEPILOGO", e["siglaEntita"], azione["SiglaAzione"], suffissoData), Tuple.Create(_rigaAttiva + i++, colonnaInizio + j));
                        }
                    }
                    j++;
                }
                colonnaInizio += _nAzioni;
            });
        }
        protected void InitBarraEntita()
        {
            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            DataView entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;

            categorie.RowFilter = "Operativa = 1";
            entita.RowFilter = "";

            object[,] values = new object[categorie.Count + entita.Count, 1];
            Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, _colonnaInizio], _ws.Cells[_rigaAttiva + values.Length - 1, _colonnaInizio]];
            rng.Style = "recapEntityBarStyle";
            Style.RangeStyle(rng, "borders:[top:medium,right:medium,bottom:medium,left:medium,insideh:thin]");
            int i = 0;
            foreach (DataRowView categoria in categorie)
            {
                Excel.Range titoloCategoria = _ws.Range[_ws.Cells[_rigaAttiva + i, _colonnaInizio], _ws.Cells[_rigaAttiva + i, _colonnaInizio + ((Struct.intervalloGiorni + 1) * _nAzioni)]];
                titoloCategoria.Merge();
                titoloCategoria.Style = "recapCategoryTitle";
                Style.RangeStyle(titoloCategoria, "Borders:[left:medium, top:medium, right:medium]");

                values[i++, 0] = categoria["DesCategoria"];
                entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "'";

                

                foreach (DataRowView e in entita)
                {
                    _nomiDefiniti.Add(DefinedNames.GetName("RIEPILOGO", e["siglaEntita"], "GOTO"), Tuple.Create(_rigaAttiva + i, _colonnaInizio));
                    values[i++, 0] = (e["Gerarchia"] is DBNull ? "" : "     ") + e["DesEntita"];
                }
            }
            rng.Value = values;

            categorie.RowFilter = "";
            entita.RowFilter = "";
            rng.EntireColumn.AutoFit();
        }
        protected void AbilitaAzioni()
        {
            DataView entitaAzioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAAZIONE].DefaultView;
            entitaAzioni.RowFilter = "";

            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                foreach (DataRowView entitaAzione in entitaAzioni)
                {
                    string nome = DefinedNames.GetName("RIEPILOGO", entitaAzione["SiglaEntita"], entitaAzione["SiglaAzione"]);
                    if (_nomiDefiniti.IsDefined(nome))
                    {
                        Tuple<int, int>[] celleAzione = _nomiDefiniti[nome];

                        foreach (Tuple<int, int> cella in celleAzione)
                        {
                            Style.RangeStyle(_ws.Cells[cella.Item1, cella.Item2], "BackPattern: none");
                        }
                    }
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
                        string nome = DefinedNames.GetName("RIEPILOGO", valore["SiglaEntita"], valore["SiglaAzione"], suffissoData);
                        if (_nomiDefiniti.IsDefined(nome))
                        {
                            Tuple<int, int> cella = _nomiDefiniti[nome][0];
                            string commento = "";

                            Excel.Range rng = _ws.Cells[cella.Item1, cella.Item2];

                            if (valore["Presente"].Equals("1"))
                            {
                                rng.ClearComments();
                                DateTime data = DateTime.ParseExact(valore["Data"].ToString(), "yyyyMMddHHmm", CultureInfo.InvariantCulture);
                                commento = "Utente: " + valore["Utente"] + "\nData: " + data.ToString("dd MMM yyyy") + "\nOra: " + data.ToString("HH:mm");
                                rng.AddComment(commento);
                                rng.Value = "OK";
                                Style.RangeStyle(rng, "BackColor:4;Align:Center");
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
