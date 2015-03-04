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
using Iren.ToolsExcel.Core;

namespace Iren.ToolsExcel.Base
{
    public class Riepilogo: CommonFunctions
    {
        #region Variabili
        
        Excel.Worksheet _ws;
        DefinedNames _nomiDefiniti;
        Cell _cell;
        Struttura _struttura;
        int _rigaAttiva;
        int _colonnaInizio;
        int _nAzioni;
        static bool _resizeFatto = false;

        #endregion

        #region Costruttori

        public Riepilogo(Excel.Worksheet ws)
        {
            _ws = ws;

            //dimensionamento celle in base ai parametri del DB
            DataView paramApplicazione = LocalDB.Tables[Tab.APPLICAZIONE].DefaultView;

            _cell = new Cell();
            _struttura = new Struttura();

            //prendo i valori di default
            _cell.Width.empty = double.Parse(paramApplicazione[0]["ColVuotaWidth"].ToString());
            _cell.Width.dato = double.Parse(paramApplicazione[0]["ColDatoWidth"].ToString());
            _cell.Width.entita = double.Parse(paramApplicazione[0]["ColEntitaWidth"].ToString());
            _cell.Width.informazione = double.Parse(paramApplicazione[0]["ColInformazioneWidth"].ToString());
            _cell.Width.unitaMisura = double.Parse(paramApplicazione[0]["ColUMWidth"].ToString());
            _cell.Width.parametro = double.Parse(paramApplicazione[0]["ColParametroWidth"].ToString());
            _cell.Height.normal = double.Parse(paramApplicazione[0]["RowHeight"].ToString());
            _cell.Height.empty = double.Parse(paramApplicazione[0]["RowVuotaHeight"].ToString());
            
            _struttura.rigaBlock = 5;
            _struttura.colBlock = 59;

            _nomiDefiniti = new DefinedNames(_ws.Name);
        }

        #endregion

        #region Metodi

        private void CicloGiorni(Action<int, string, DateTime> callback)
        {
            DateTime dataInizio = CommonFunctions.DB.DataAttiva;
            DateTime dataFine = CommonFunctions.DB.DataAttiva.AddDays(Simboli.intervalloGiorni);
            CicloGiorni(dataInizio, dataFine, callback);
        }
        private void CicloGiorni(DateTime dataInizio, DateTime dataFine, Action<int, string, DateTime> callback)
        {
            for (DateTime giorno = dataInizio; giorno <= dataFine; giorno = giorno.AddDays(1))
            {
                int oreGiorno = GetOreGiorno(giorno);
                string suffissoData = GetSuffissoData(dataInizio, giorno);

                if (giorno == dataInizio && _struttura.visData0H24)
                {
                    oreGiorno++;
                }

                callback(oreGiorno, suffissoData, giorno);
            }
        }

        public void InitLabels()
        {
            //inizializzo i label
            _ws.Shapes.Item("lbTitolo").TextFrame.Characters().Text = Simboli.nomeApplicazione;
            _ws.Shapes.Item("lbDataInizio").TextFrame.Characters().Text = CommonFunctions.DB.DataAttiva.ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbDataFine").TextFrame.Characters().Text = CommonFunctions.DB.DataAttiva.AddDays(Simboli.intervalloGiorni).ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbVersione").TextFrame.Characters().Text = "Foglio v." + WorkbookVersion.ToString();
            _ws.Shapes.Item("lbUtente").TextFrame.Characters().Text = "Utente: " + LocalDB.Tables[Tab.UTENTE].Rows[0]["Nome"];

            //aggiorna la scritta di modifica dati
            Simboli.ModificaDati = false;

            //aggiorna la scritta e il colore del label che mostra l'ambiente
            Simboli.Ambiente = ConfigurationManager.AppSettings["DB"];

            if (Simboli.intervalloGiorni > 0)
            {
                if (!_resizeFatto)
                {
                    _ws.Shapes.Item("lbDataInizio").ScaleWidth(0.4819f, Office.MsoTriState.msoFalse);
                    _ws.Shapes.Item("lbDataFine").Visible = Office.MsoTriState.msoTrue;
                    _resizeFatto = true;
                }
            }
            else
            {
                _ws.Shapes.Item("lbDataInizio").Width = 485.8582677165f;
                _ws.Shapes.Item("lbDataFine").Visible = Office.MsoTriState.msoFalse;
            }
        }        
        private void Clear()
        {
            int dataOreTot = GetOreIntervallo(CommonFunctions.DB.DataAttiva, CommonFunctions.DB.DataAttiva.AddDays(Simboli.intervalloGiorni)) + (_struttura.visData0H24 ? 1 : 0) + (_struttura.visParametro ? 1 : 0);

            _ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            _ws.UsedRange.EntireColumn.Delete();
            _ws.UsedRange.FormatConditions.Delete();
            _ws.UsedRange.EntireRow.Hidden = false;
            _ws.UsedRange.Font.Size = 8;
            _ws.UsedRange.NumberFormat = "General";
            _ws.UsedRange.Font.Name = "Verdana";

            _ws.Range[_ws.Cells[1, 1], _ws.Cells[1, _struttura.colRecap - 1]].EntireColumn.ColumnWidth = _cell.Width.empty;            
            _ws.Rows[1].RowHeight = _cell.Height.empty;

            ((Excel._Worksheet)_ws).Activate();
            _ws.Application.ActiveWindow.FreezePanes = false;
            _ws.Cells[_struttura.rigaBlock, _struttura.colBlock].Select();
            _ws.Application.ActiveWindow.ScrollColumn = 1;
            _ws.Application.ActiveWindow.ScrollRow = 1;
            _ws.Application.ActiveWindow.FreezePanes = true;
        }

        public void LoadStructure()
        {
            _colonnaInizio = _struttura.colRecap;
            _rigaAttiva = _struttura.rowRecap;

            InitLabels();
            Clear();

            DataView azioni = LocalDB.Tables[Tab.AZIONE].DefaultView;
            DataView categorie = LocalDB.Tables[Tab.CATEGORIA].DefaultView;
            DataView entita = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;

            categorie.RowFilter = "Operativa = 1";
            azioni.RowFilter = "Visibile = 1 AND Operativa = 1";
            entita.RowFilter = "";

            InitBarraTitolo(azioni);
            _rigaAttiva += 3;
            FormattaAllDati(azioni, categorie, entita);
            InitBarraEntita(categorie, entita);
            AbilitaAzioni();
            CaricaDatiRiepilogo();

            //Se sono in multiscreen lascio il riepilogo alla fine, altrimenti lo riporto all'inizio
            if (Screen.AllScreens.Length == 1)
            {
                _ws.Application.ActiveWindow.SmallScroll(Type.Missing, Type.Missing, _struttura.colRecap - _struttura.colBlock - 1);
            }
        }

        private void InitBarraTitolo(DataView azioni)
        {
            _nAzioni = 0;

            Dictionary<object, List<object>> valAzioni = new Dictionary<object, List<object>>();
            foreach (DataRowView azione in azioni)
            {
                if (!valAzioni.ContainsKey(azione["Gerarchia"]))
                {
                    valAzioni.Add(azione["Gerarchia"], new List<object>() { azione["DesAzioneBreve"] });
                }
                else
                {
                    valAzioni[azione["Gerarchia"]].Add(azione["DesAzioneBreve"]);
                }
                _nAzioni++;
            }            
            int nAzioniPadre = valAzioni.Count;

            //numero totale di celle della barra del titolo
            object[,] values = new object[3, _nAzioni];
            //la prima libera per mettere la data successivamente
            int[] azioniPerPadre = new int[valAzioni.Count];
            int ipadre = 0;
            int iazioni = 0;
            int j = 0;
            foreach (KeyValuePair<object, List<object>> keyVal in valAzioni)
            {
                azioniPerPadre[j++] = keyVal.Value.Count;
                values[1, ipadre] = keyVal.Key.ToString().ToUpperInvariant();
                ipadre += keyVal.Value.Count;
                foreach (object nomeAzione in keyVal.Value) 
                    values[2, iazioni++] = nomeAzione.ToString().ToUpperInvariant();
            }

            int colonnaInizio = _colonnaInizio + 1;
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                values[0, 0] = giorno;
                Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, colonnaInizio], _ws.Cells[_rigaAttiva + 2, colonnaInizio + _nAzioni - 1]];

                _nomiDefiniti.Add(DefinedNames.GetName("RIEPILOGO", "T", suffissoData), _rigaAttiva, colonnaInizio, _rigaAttiva, colonnaInizio + _nAzioni - 1);

                rng.Style = "recapTitleBarStyle";
                rng.Rows[1].Merge();
                Style.RangeStyle(rng.Rows[1], "FontSize:10;NumberFormat:[ddd d mmm yyyy]");
                int i = 1;
                foreach (int numAzioniPerPadre in azioniPerPadre)
                {
                    _ws.Range[rng.Cells[2, i], rng.Cells[2, i + numAzioniPerPadre - 1]].Merge();
                    Style.RangeStyle(_ws.Range[rng.Cells[3, i], rng.Cells[3, i + numAzioniPerPadre - 1]], "FontSize:7;Borders:[left:medium, bottom:medium, right:medium, insidev:thin]");
                    i += numAzioniPerPadre;
                }
                rng.Value = values;

                colonnaInizio += _nAzioni;
            });
        }
        private void FormattaAllDati(DataView azioni, DataView categorie, DataView entita)
        {
            azioni.RowFilter = "Gerarchia IS NOT NULL";
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

            azioni.RowFilter = "";
            entita.RowFilter = "";
        }
        private void InitBarraEntita(DataView categorie, DataView entita)
        {
            object[,] values = new object[categorie.Count + entita.Count, 1];
            Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, _colonnaInizio], _ws.Cells[_rigaAttiva + values.Length - 1, _colonnaInizio]];
            rng.Style = "recapEntityBarStyle";
            Style.RangeStyle(rng, "borders:[top:medium,right:medium,bottom:medium,left:medium,insideh:thin]");
            int i = 0;
            foreach (DataRowView categoria in categorie)
            {
                Excel.Range titoloCategoria = _ws.Range[_ws.Cells[_rigaAttiva + i, _colonnaInizio], _ws.Cells[_rigaAttiva + i, _colonnaInizio + ((Simboli.intervalloGiorni + 1) * _nAzioni)]];
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
        private void AbilitaAzioni()
        {
            DataView entitaAzioni = LocalDB.Tables[Tab.ENTITAAZIONE].DefaultView;
            entitaAzioni.RowFilter = "";

            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                foreach (DataRowView entitaAzione in entitaAzioni)
                {
                    string nome = DefinedNames.GetName("RIEPILOGO", entitaAzione["SiglaEntita"], entitaAzione["SiglaAzione"]);
                    Tuple<int, int>[] celleAzione = _nomiDefiniti[nome];

                    foreach (Tuple<int, int> cella in celleAzione)
                    {
                        Style.RangeStyle(_ws.Cells[cella.Item1, cella.Item2], "BackPattern: none");
                    }
                }
            });
        }
        private void CaricaDatiRiepilogo()
        {
            try
            {
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    DataView datiRiepilogo = DB.Select(DataBase.SP.APPLICAZIONE_RIEPILOGO, "@Data=" + giorno.ToString("yyyyMMdd")).DefaultView;
                    foreach (DataRowView valore in datiRiepilogo)
                    {
                        string nome = DefinedNames.GetName("RIEPILOGO", valore["SiglaEntita"], valore["SiglaAzione"], suffissoData);
                        Tuple<int, int> cella = _nomiDefiniti[nome][0];
                        string commento = "";

                        Excel.Range rng = _ws.Cells[cella.Item1, cella.Item2];

                        if (valore["Presente"].Equals("1"))
                        {
                            DateTime data = DateTime.ParseExact(valore["Data"].ToString(), "yyyyMMddHHmm", CultureInfo.InvariantCulture);
                            commento = "Utente: " + valore["Utente"] + "\nData: " + data.ToString("dd MMM yyyy") + "\nOra: " + data.ToString("HH:mm");
                            rng.AddComment(commento);
                            rng.Value = "OK";
                            Style.RangeStyle(rng, "BackColor:4;Align:Center");
                        }
                    }
                });
            }
            catch (Exception e)
            {
                CommonFunctions.InsertLog(DataBase.TipologiaLOG.LogErrore, "CaricaDatiRiepilogo: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void AggiornaRiepilogo(object entita, object azione, bool presente, DateTime? dataRif = null)
        {
            if(dataRif == null)
                dataRif = CommonFunctions.DB.DataAttiva;

            Tuple<int, int> cella = _nomiDefiniti[DefinedNames.GetName("RIEPILOGO", entita, azione, GetSuffissoData(CommonFunctions.DB.DataAttiva, dataRif.Value))][0];
            Excel.Range rng = _ws.Cells[cella.Item1, cella.Item2];

            if (presente)
            {
                string commento = "Utente: " + LocalDB.Tables[Tab.UTENTE].Rows[0]["Nome"] + "\nData: " + DateTime.Now.ToString("dd MMM yyyy") + "\nOra: " + DateTime.Now.ToString("HH:mm");
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

        private void AggiornaDate()
        {
            _ws.Shapes.Item("lbDataInizio").TextFrame.Characters().Text = CommonFunctions.DB.DataAttiva.ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbDataFine").TextFrame.Characters().Text = CommonFunctions.DB.DataAttiva.AddDays(Simboli.intervalloGiorni).ToString("ddd d MMM yyyy");

            CicloGiorni((oreGiorno, suffissoData, giorno) => 
            {
                Tuple<int, int>[] riga = _nomiDefiniti.GetRange(DefinedNames.GetName("RIEPILOGO", "T", suffissoData));
                _ws.Range[_ws.Cells[riga[0].Item1, riga[0].Item2], _ws.Cells[riga[1].Item1, riga[1].Item2]].Value = giorno;
            });
        }
        public void UpdateRiepilogo()
        {
            AggiornaDate();
            AbilitaAzioni();
            CaricaDatiRiepilogo();
        }

        private void DisabilitaTutto()
        {
            DataView categorie = LocalDB.Tables[Tab.CATEGORIA].DefaultView;
            DataView entita = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;

            categorie.RowFilter = "Operativa = 1";

            foreach (DataRowView categoria in categorie)
            {
                entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "'";
                
                foreach (DataRowView e in entita)
                {
                    Tuple<int, int>[] riga = _nomiDefiniti.Get(DefinedNames.GetName("RIEPILOGO", e["siglaEntita"]), "GOTO");

                    Excel.Range rng = _ws.Range[_ws.Cells[riga[0].Item1, riga[0].Item2], _ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]];
                    rng.Value = "";
                    rng.ClearComments();

                    Style.RangeStyle(rng, "BackPattern: CrissCross");
                }
            }

        }
        public void RiepilogoInEmergenza()
        {
            AggiornaDate();
            DisabilitaTutto();
        }

        #endregion

    }
}
