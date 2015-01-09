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

        private void CicloGiorni(Func<int, string, DateTime, bool> callback)
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

            if (true)//_struttura.intervalloGiorni > 0)
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

            InitBarraTitolo();
            InitBarraEntita();
        }

        private void InitBarraTitolo()
        {
            DataView azioni = LocalDB.Tables[Tab.AZIONE].DefaultView;
            int nAzioni = 0;

            Dictionary<object, List<object>> valAzioni = new Dictionary<object, List<object>>();
            Dictionary<object, object> valAzioniPadre = new Dictionary<object, object>();
            foreach (DataRowView azione in azioni)
            {
                if (azione["Gerarchia"] is DBNull) 
                {
                    valAzioni.Add(azione["DesAzioneBreve"], new List<object>());
                    valAzioniPadre.Add(azione["SiglaAzione"], azione["DesAzioneBreve"]);
                }
                else
                    if (!valAzioniPadre.ContainsKey(azione["Gerarchia"])) 
                    {
                        valAzioni.Add(azione["DesAzioneBreve"], new List<object>());
                        valAzioniPadre.Add(azione["SiglaAzione"], azione["DesAzioneBreve"]);
                    }
                    else
                    {
                        valAzioni[valAzioniPadre[azione["Gerarchia"]]].Add(azione["DesAzioneBreve"]);
                        nAzioni++;
                    }
            }            
            int nAzioniPadre = valAzioni.Count;

            //numero totale di celle della barra del titolo
            object[,] values = new object[3, nAzioni];
            //la prima libera per mettere la data successivamente
            int[] azioniPerPadre = new int[valAzioni.Count];
            int ipadre = 0;
            int iazioni = 0;
            int j = 0;
            foreach (KeyValuePair<object, List<object>> keyVal in valAzioni)
            {
                azioniPerPadre[j++] = keyVal.Value.Count;
                values[1, ipadre] = keyVal.Key.ToString().ToUpperInvariant();
                ipadre += 2;
                foreach (object nomeAzione in keyVal.Value) 
                    values[2, iazioni++] = nomeAzione.ToString().ToUpperInvariant();
            }

            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                values[0, 0] = giorno;
                Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, _colonnaInizio + 1], _ws.Cells[_rigaAttiva + 2, _colonnaInizio + nAzioni]];
                rng.Style = "recapTitleBarStyle";
                rng.Rows[1].Merge();
                Style.RangeStyle(rng.Rows[1], "FontSize:10;NumberFormat:[ddd d mmm yyyy]");
                int colonnaInizio = 1;
                foreach (int numAzioni in azioniPerPadre)
                {
                    _ws.Range[rng.Cells[2, colonnaInizio], rng.Cells[2, colonnaInizio + numAzioni - 1]].Merge();
                    Style.RangeStyle(_ws.Range[rng.Cells[3, colonnaInizio], rng.Cells[3, colonnaInizio + numAzioni - 1]], "FontSize:7;Borders:[left:medium, bottom:medium, right:medium, insidev:thin]");
                    colonnaInizio += numAzioni;
                }
                rng.Value = values;

                return true;
            });
        }
        private void InitBarraEntita()
        {
            DataView categorie = LocalDB.Tables[Tab.CATEGORIA].DefaultView;
            DataView entita = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;
            categorie.RowFilter = "Operativa = 1";

            object[,] values = new object[categorie.Count + entita.Count, 1];
            Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva + 3, _colonnaInizio], _ws.Cells[_rigaAttiva + 3 + values.Length - 1, _colonnaInizio]];
            rng.Style = "recapEntityBarStyle";
            Style.RangeStyle(rng, "borders:[top:medium,right:medium,bottom:medium,left:medium,insideh:thin]");
            int i = 0;
            foreach (DataRowView categoria in categorie)
            {
                values[i++, 0] = categoria["DesCategoria"];
                entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "'";
                foreach (DataRowView ent in entita)
                {
                    values[i++, 0] = (ent["Gerarchia"] is DBNull ? "" : "     ") + ent["DesEntita"];
                }
            }
            rng.Value = values;

            categorie.RowFilter = "";
            entita.RowFilter = "";
            rng.EntireColumn.AutoFit();
        }


        #endregion

    }
}
