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

namespace Iren.FrontOffice.Base
{
    public class Riepilogo: CommonFunctions
    {
        #region Variabili
        
        Excel.Worksheet _ws;
        Dictionary<string, object> _config = new Dictionary<string,object>();
        DateTime _dataInizio;
        DateTime _dataFine;
        DefinedNames _nomiDefiniti;
        Cell _cell;
        Struttura _struttura;
        int _rigaAttiva;
        int _colonnaInizio;
        int _nAzioni;        


        #endregion

        #region Costruttori

        public Riepilogo(Excel.Worksheet ws)
        {
            _ws = ws;

            _config.Add("DataInizio", DateTime.ParseExact(ConfigurationManager.AppSettings["DataInizio"], "yyyyMMdd", CultureInfo.InvariantCulture));

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
            _struttura.intervalloGiorni = (int)paramApplicazione[0]["IntervalloGiorni"];
            _struttura.colBlock = 59;

            _nomiDefiniti = new DefinedNames(_ws.Name);
        }

        #endregion

        #region Metodi

        private void CicloGiorni(Action<int, string, DateTime> callback)
        {
            for (DateTime giorno = _dataInizio; giorno <= _dataFine; giorno = giorno.AddDays(1))
            {
                int oreGiorno = GetOreGiorno(giorno);
                string suffissoData = GetSuffissoData(_dataInizio, giorno);

                if (giorno == _dataInizio && _struttura.visData0H24)
                {
                    oreGiorno++;
                }

                callback(oreGiorno, suffissoData, giorno);
            }
        }

        private void Clear()
        {
            //inizializzo i label
            _ws.Shapes.Item("lbTitolo").TextFrame.Characters().Text = Simboli.nomeApplicazione;
            _ws.Shapes.Item("lbDataInizio").TextFrame.Characters().Text = _dataInizio.ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbDataFine").TextFrame.Characters().Text = _dataFine.ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbVersione").TextFrame.Characters().Text = "Foglio v." + WorkbookVersion.ToString();
            _ws.Shapes.Item("lbUtente").TextFrame.Characters().Text = "Utente: " + LocalDB.Tables[Tab.UTENTE].Rows[0]["Nome"];
            //TODO controllo DB e stati modifica/ambiente


            if (_struttura.intervalloGiorni > 0)
            {
                _ws.Shapes.Item("lbDataInizio").ScaleWidth(0.4819f, Office.MsoTriState.msoFalse);
                _ws.Shapes.Item("lbDataFine").Visible = Office.MsoTriState.msoTrue;
            }
            else
            {
                _ws.Shapes.Item("lbDataInizio").ScaleWidth(1f, Office.MsoTriState.msoFalse);
                _ws.Shapes.Item("lbDataFine").Visible = Office.MsoTriState.msoFalse;
            }

            int dataOreTot = GetOreIntervallo(_dataInizio, _dataInizio.AddDays(_struttura.intervalloGiorni)) + (_struttura.visData0H24 ? 1 : 0) + (_struttura.visParametro ? 1 : 0);

            _ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            _ws.UsedRange.EntireColumn.Delete();
            _ws.UsedRange.FormatConditions.Delete();
            _ws.UsedRange.EntireRow.Hidden = false;
            _ws.UsedRange.Font.Size = 8;
            _ws.UsedRange.NumberFormat = "General";
            _ws.UsedRange.Font.Name = "Verdana";

            _ws.Range[_ws.Cells[1, 1], _ws.Cells[1, _struttura.colRecap - 1]].EntireColumn.ColumnWidth = _cell.Width.empty;            
            _ws.Rows[1].RowHeight = _cell.Height.empty;

            _ws.Activate();
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
            _dataInizio = (DateTime)_config["DataInizio"];
            _dataFine = _dataInizio.AddDays(_struttura.intervalloGiorni);

            Clear();

            DataView azioni = LocalDB.Tables[Tab.AZIONE].DefaultView;
            DataView categorie = LocalDB.Tables[Tab.CATEGORIA].DefaultView;
            DataView entita = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;
            DataView entitaAzioni = LocalDB.Tables[Tab.ENTITAAZIONE].DefaultView;

            categorie.RowFilter = "Operativa = 1";
            azioni.RowFilter = "Visibile = 1 AND Operativa = 1";
            entita.RowFilter = "";
            entitaAzioni.RowFilter = "";

            InitBarraTitolo(azioni);
            _rigaAttiva += 3;
            FormattaAllDati(azioni, categorie, entita);
            InitBarraEntita(categorie, entita);
            AbilitaAzioni(entitaAzioni);
            CaricaDatiRiepilogo();
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
                            _nomiDefiniti.Add(GetName("RIEPILOGO", e["siglaEntita"], azione["SiglaAzione"], suffissoData), Tuple.Create(_rigaAttiva + i++, colonnaInizio + j));
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
                Excel.Range titoloCategoria = _ws.Range[_ws.Cells[_rigaAttiva + i, _colonnaInizio], _ws.Cells[_rigaAttiva + i, _colonnaInizio + ((_struttura.intervalloGiorni + 1) * _nAzioni)]];
                titoloCategoria.Merge();
                titoloCategoria.Style = "recapCategoryTitle";
                Style.RangeStyle(titoloCategoria, "Borders:[left:medium, top:medium, right:medium]");

                values[i++, 0] = categoria["DesCategoria"];
                entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "'";

                

                foreach (DataRowView e in entita)
                {
                    _nomiDefiniti.Add(GetName("RIEPILOGO", e["siglaEntita"], "GOTO"), Tuple.Create(_rigaAttiva + i, _colonnaInizio));
                    values[i++, 0] = (e["Gerarchia"] is DBNull ? "" : "     ") + e["DesEntita"];
                }
            }
            rng.Value = values;

            categorie.RowFilter = "";
            entita.RowFilter = "";
            rng.EntireColumn.AutoFit();
        }
        private void AbilitaAzioni(DataView entitaAzioni)
        {
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                foreach (DataRowView entitaAzione in entitaAzioni)
                {
                    string nome = GetName("RIEPILOGO", entitaAzione["SiglaEntita"], entitaAzione["SiglaAzione"]);
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
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                DataView datiRiepilogo = DB.Select("spApplicazioneRiepilogo", "@Data=" + giorno.ToString("yyyyMMdd")).DefaultView;                
                foreach (DataRowView valore in datiRiepilogo)
                {
                    string nome = GetName("RIEPILOGO", valore["SiglaEntita"], valore["SiglaAzione"], suffissoData);
                    Tuple<int, int> cella = _nomiDefiniti[nome][0];
                    string commento = "";

                    Excel.Range rng = _ws.Cells[cella.Item1, cella.Item2];

                    if(valore["Presente"].Equals("1")) 
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

        #endregion

    }
}
