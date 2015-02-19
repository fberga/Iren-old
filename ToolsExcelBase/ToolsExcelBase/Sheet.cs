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
using System.Diagnostics;
using System.Drawing;

namespace Iren.FrontOffice.Base
{
    public class Sheet : CommonFunctions, IDisposable
    {
        #region Variabili

        Excel.Worksheet _ws;
        Dictionary<string, object> _config = new Dictionary<string, object>();
        DateTime _dataInizio;
        DateTime _dataFine;
        int _colonnaInizio;
        int _intervalloOre;
        int _rigaAttiva;
        bool _disposed = false;
        DefinedNames _nomiDefiniti;
        Cell _cell;
        Struttura _struttura;

        #endregion

        #region Costruttori

        public Sheet(Excel.Worksheet ws)
        {
            _ws = ws;

            DataView categorie = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.CATEGORIA].DefaultView;
            categorie.RowFilter = "DesCategoria = '" + ws.Name + "'";

            _config.Add("SiglaCategoria", categorie[0]["SiglaCategoria"]);
            _config.Add("DataInizio", DateTime.ParseExact(ConfigurationManager.AppSettings["DataInizio"], "yyyyMMdd", CultureInfo.InvariantCulture));

            //dimensionamento celle in base ai parametri del DB
            DataView paramApplicazione = LocalDB.Tables[Tab.APPLICAZIONE].DefaultView;

            _cell = new Cell();
            _struttura = new Struttura();

            _cell.Width.empty = double.Parse(paramApplicazione[0]["ColVuotaWidth"].ToString());
            _cell.Width.dato = double.Parse(paramApplicazione[0]["ColDatoWidth"].ToString());
            _cell.Width.entita = double.Parse(paramApplicazione[0]["ColEntitaWidth"].ToString());
            _cell.Width.informazione = double.Parse(paramApplicazione[0]["ColInformazioneWidth"].ToString());
            _cell.Width.unitaMisura = double.Parse(paramApplicazione[0]["ColUMWidth"].ToString());
            _cell.Width.parametro = double.Parse(paramApplicazione[0]["ColParametroWidth"].ToString());
            _cell.Height.normal = double.Parse(paramApplicazione[0]["RowHeight"].ToString());
            _cell.Height.empty = double.Parse(paramApplicazione[0]["RowVuotaHeight"].ToString());
            _struttura.rigaBlock = (int)paramApplicazione[0]["RowBlocco"];
            _struttura.rigaGoto = (int)paramApplicazione[0]["RowGoto"];
            _struttura.intervalloGiorni = (int)paramApplicazione[0]["IntervalloGiorni"];
            _struttura.visData0H24 = paramApplicazione[0]["VisData0H24"].ToString() == "1";
            _struttura.visParametro = paramApplicazione[0]["VisParametro"].ToString() == "1";
            _struttura.colBlock = (int)paramApplicazione[0]["ColBlocco"] + (_struttura.visParametro ? 1 : 0);

            _nomiDefiniti = new DefinedNames(_ws.Name);
        }
        ~Sheet()
        {
            Dispose();
        }

        #endregion

        #region Parametri

        private int VisParametro
        {
            get
            {
                return _struttura.visParametro ? 3 : 2;
            }
        }

        #endregion

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
            int dataOreTot = GetOreIntervallo(_dataInizio, _dataInizio.AddDays(_struttura.intervalloGiorni)) + (_struttura.visData0H24 ? 1 : 0) + (_struttura.visParametro ? 1 : 0);

            _ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            _ws.UsedRange.EntireColumn.Delete();
            _ws.UsedRange.FormatConditions.Delete();
            _ws.UsedRange.EntireRow.Hidden = false;
            _ws.UsedRange.Font.Size = 10;
            _ws.UsedRange.NumberFormat = "General";
            _ws.UsedRange.Font.Name = "Verdana";
            _ws.UsedRange.RowHeight = _cell.Height.normal;

            _ws.Rows["1:" + (_struttura.rigaBlock - 1)].RowHeight = _cell.Height.empty;
            _ws.Rows[_struttura.rigaGoto].RowHeight = _cell.Height.normal;

            _ws.Columns[1].ColumnWidth = _cell.Width.empty;
            _ws.Columns[2].ColumnWidth = _cell.Width.entita;

            _ws.Activate();
            _ws.Application.ActiveWindow.FreezePanes = false;
            _ws.Cells[_struttura.rigaBlock, _struttura.colBlock].Select();
            _ws.Application.ActiveWindow.ScrollColumn = 1;
            _ws.Application.ActiveWindow.ScrollRow = 1;
            _ws.Application.ActiveWindow.FreezePanes = true;

            string gotoBarRangeName = DefinedNames.GetName(_config["SiglaCategoria"], "GOTO_BAR");
            Excel.Range rng = _ws.Range[_ws.Cells[2, 2], _ws.Cells[_struttura.rigaBlock - 2, _struttura.colBlock + dataOreTot - 1]];
            rng.Style = "gotoBarStyle";
            rng.BorderAround2(Weight: Excel.XlBorderWeight.xlMedium, Color: 1);

            int infoCols = _struttura.colBlock - VisParametro;

            _ws.Columns[infoCols].ColumnWidth = _cell.Width.informazione;
            _ws.Columns[infoCols + 1].ColumnWidth = _cell.Width.unitaMisura;
            if (_struttura.visParametro)
                _ws.Columns[infoCols + 2].ColumnWidth = _cell.Width.parametro;
        }

        public void LoadStructure()
        {
            Stopwatch watch = Stopwatch.StartNew();

            DataView dvEP = LocalDB.Tables[Tab.ENTITAPROPRIETA].DefaultView;
            DataView dvCE = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;

            dvCE.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";
            _dataInizio = (DateTime)_config["DataInizio"];

            Clear();
            InitBarraNavigazione(dvCE);

            _rigaAttiva = _struttura.rigaBlock;

            foreach (DataRowView entita in dvCE)
            {
                string siglaEntita = "" + entita["SiglaEntita"];
                dvEP.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                if (dvEP.Count > 0)
                    _dataFine = _dataInizio.AddDays(double.Parse("" + dvEP[0]["Valore"]));
                else
                    _dataFine = _dataInizio.AddDays(_struttura.intervalloGiorni);

                InitBloccoEntita(entita);
            }
            dvEP.RowFilter = "";
            dvCE.RowFilter = "";
            watch.Stop();

            watch = Stopwatch.StartNew();
            CaricaInformazioni();
            watch.Stop();
            watch = Stopwatch.StartNew();
            CalcolaFormule();
            watch.Stop();
            watch = Stopwatch.StartNew();
            AggiornaFormule(_ws);
            watch.Stop();
            watch = Stopwatch.StartNew();
            InsertGrafici();
            watch.Stop();
        }

        public void UpdateData()
        {
            CancellaDati();
            AggiornaDateTitoli();
            CaricaParametri();
            CaricaInformazioni();
            AggiornaGrafici();
        }

        private void InitBarraNavigazione(DataView entita)
        {
            object[] descrizioni = new object[entita.Count];
            int i = -1;
            foreach (DataRowView e in entita)
            {
                descrizioni[++i] = e["DesEntitaBreve"];
                _nomiDefiniti.Add(DefinedNames.GetName(e["siglaEntita"], "GOTO"), Tuple.Create(_struttura.rigaGoto, _struttura.colBlock + i));
            }

            Excel.Range rng = _ws.Range[_ws.Cells[_struttura.rigaGoto, _struttura.colBlock],
                _ws.Cells[_struttura.rigaGoto, _struttura.colBlock + i]];

            rng.Value = descrizioni;
            rng.Style = "navBarStyle";
        }

        private void InitBloccoEntita(DataRowView entita)
        {
            _rigaAttiva++;
            DataView grafici = LocalDB.Tables[Tab.ENTITAGRAFICO].DefaultView;
            grafici.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

            DataView graficiInfo = LocalDB.Tables[Tab.ENTITAGRAFICOINFORMAZIONE].DefaultView;
            graficiInfo.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

            DataView informazioni = LocalDB.Tables[Tab.ENTITAINFORMAZIONE].DefaultView;
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
            informazioni.Sort = "Ordine";

            DataView parametriD = LocalDB.Tables[Tab.ENTITAPARAMETROD].DefaultView;
            DataView parametriH = LocalDB.Tables[Tab.ENTITAPARAMETROH].DefaultView;

            DataView formattazione = LocalDB.Tables[Tab.ENTITAINFORMAZIONEFORMATTAZIONE].DefaultView;

            _colonnaInizio = _struttura.colBlock;
            _intervalloOre = GetOreIntervallo(_dataInizio, _dataFine) + (_struttura.visData0H24 ? 1 : 0) + (_struttura.visParametro ? 1 : 0);

            //titolo + data
            //Stopwatch watch = Stopwatch.StartNew();
            InsertTitoloEntita(entita);
            //watch.Stop();
            //creazione range grafici
            //watch = Stopwatch.StartNew();
            InsertRangeGrafici(grafici);
            //watch.Stop();
            //ore
            //watch = Stopwatch.StartNew();
            InsertOre(entita["SiglaEntita"]);
            //watch.Stop();
            //titolo verticale
            //watch = Stopwatch.StartNew();
            InsertTitoloVerticale(entita["SiglaEntita"], entita["DesEntitaBreve"], informazioni.Count);
            //watch.Stop();
            //formatta AllDati
            //watch = Stopwatch.StartNew();
            FormattaAllDati(informazioni);
            //watch.Stop();
            //informazioni
            //watch = Stopwatch.StartNew();
            InsertInformazioniEntita(entita["SiglaEntita"], informazioni);
            //watch.Stop();
            //nomi
            //watch = Stopwatch.StartNew();
            CreaNomiCelle(informazioni);
            //watch.Stop();
            //valori e formule
            //watch = Stopwatch.StartNew();
            informazioni.RowFilter += " AND (ValoreDefault IS NOT NULL OR FormulaInCella = 1)";
            InsertValoriCelle(informazioni);
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
            //watch.Stop();
            //parametri giornalieri/orari
            //watch = Stopwatch.StartNew();
            informazioni.RowFilter += " AND SiglaTipologiaParametro IS NOT NULL";
            InsertParametri(informazioni, parametriD, parametriH);
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
            //watch.Stop();
            //formattazione condizionale
            //watch = Stopwatch.StartNew();
            CreaFormattazioneCondizionale(informazioni, formattazione);
            //watch.Stop();

            //due righe vuote tra un'entità e la successiva
            _rigaAttiva += informazioni.Count + 1;
        }
        #region Blocco entità

        private void InsertTitoloEntita(DataRowView entita)
        {
            int colonnaInizio = _colonnaInizio;

            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                bool isVisibleData0H24 = giorno == _dataInizio && _struttura.visData0H24;

                if (isVisibleData0H24)
                {
                    colonnaInizio++;
                    oreGiorno--;
                }

                Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, colonnaInizio], _ws.Cells[_rigaAttiva, colonnaInizio + oreGiorno - 1]];

                _nomiDefiniti.Add(DefinedNames.GetName(entita["SiglaEntita"], "T", suffissoData), Tuple.Create(_rigaAttiva, colonnaInizio), Tuple.Create(_rigaAttiva, colonnaInizio + oreGiorno - 1));

                rng.Merge();
                rng.Style = "titleBarStyle";
                rng.Value = entita["DesEntita"].ToString().ToUpperInvariant();
                rng.RowHeight = 25;

                rng = _ws.Range[_ws.Cells[_rigaAttiva + 1, colonnaInizio], _ws.Cells[_rigaAttiva + 1, colonnaInizio + oreGiorno - 1]];

                _nomiDefiniti.Add(DefinedNames.GetName(entita["SiglaEntita"], suffissoData), Tuple.Create(_rigaAttiva + 1, colonnaInizio), Tuple.Create(_rigaAttiva + 1, colonnaInizio + oreGiorno - 1));

                rng.Merge();
                rng.Style = "dateBarStyle";
                rng.Value = giorno.ToString("MM/dd/yyyy");
                rng.RowHeight = 20;

                colonnaInizio += oreGiorno;
            });
            _rigaAttiva++;
        }
        private void InsertRangeGrafici(DataView grafici)
        {
            int i = 1;
            foreach (DataRowView grafico in grafici)
            {
                string graficoRange = DefinedNames.GetName(grafico["SiglaEntita"], "GRAFICO" + i++);

                Excel.Range rng = _ws.Range[_ws.Cells[++_rigaAttiva, _colonnaInizio],
                    _ws.Cells[_rigaAttiva, _colonnaInizio + _intervalloOre - 1]];

                _nomiDefiniti.Add(graficoRange, Tuple.Create(_rigaAttiva, _colonnaInizio), Tuple.Create(_rigaAttiva, _colonnaInizio + _intervalloOre - 1));
                rng.Merge();
                rng.Style = "chartsBarStyle";
                rng.RowHeight = 200;
            }
            _rigaAttiva++;
        }
        private void InsertOre(object siglaEntita)
        {
            string style = "Align=center;Bold=true;FontSize=10;ForeColor=1;BackColor=15;Format=##;Borders=[InsideV:thin,Top:medium,Bottom:medium,Right:medium,Left:medium]";
            _ws.Rows[_rigaAttiva].RowHeight = 20;
            int colonnaInizio = _colonnaInizio;

            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                Style.RangeStyle(_ws.Range[_ws.Cells[_rigaAttiva, colonnaInizio], _ws.Cells[_rigaAttiva, colonnaInizio + oreGiorno - 1]], style);
                for (int i = colonnaInizio; i < colonnaInizio + oreGiorno; i++)
                {
                    int val = i - colonnaInizio + 1;
                    if (giorno == _dataInizio && _struttura.visData0H24)
                        val = i == colonnaInizio ? 24 : i - colonnaInizio;
                    _ws.Cells[_rigaAttiva, i].Value = val;
                }
                colonnaInizio += oreGiorno;
            });
            _rigaAttiva++;
        }
        private void InsertTitoloVerticale(object siglaEntita, object siglaEntitaBreve, int numInformazioni)
        {
            int colonnaTitoloVert = _colonnaInizio - VisParametro - 1;
            Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, colonnaTitoloVert], _ws.Cells[_rigaAttiva + numInformazioni - 1, colonnaTitoloVert]];
            rng.Style = "titoloVertStyle";
            rng.Merge();
            rng.Orientation = numInformazioni == 1 ? Excel.XlOrientation.xlHorizontal : Excel.XlOrientation.xlVertical;
            rng.Font.Size = numInformazioni == 1 ? 6 : 9;
            rng.Value = siglaEntitaBreve;
        }
        private void FormattaAllDati(DataView informazioni)
        {
            int rigaAttiva = _rigaAttiva;
            int rigaInizioGruppo = rigaAttiva;
            int allDatiIndice = 1;

            bool primaRigaTitolo2 = informazioni[0]["SiglaTipologiaInformazione"].ToString() == "TITOLO2";
            int ultimaColonna = 0;
            foreach (DataRowView info in informazioni)
            {
                bool primaRiga = informazioni[0] == info;
                bool ultimaRiga = informazioni[informazioni.Count - 1] == info;

                //se non è la prima riga, se è l'ultima, se è un inizio gruppo e se prima non ho sistemato un TITOLO2, creo un range ALLDATI
                if ((!primaRiga && info["InizioGruppo"].ToString() == "1" && rigaInizioGruppo < rigaAttiva) || ultimaRiga)
                {
                    int colonnaInizioAllDati = _colonnaInizio;
                    CicloGiorni((oreGiorno, suffissoData, giorno) =>
                    {
                        Excel.Range allDati = _ws.Range[_ws.Cells[rigaInizioGruppo, colonnaInizioAllDati], _ws.Cells[rigaAttiva - (ultimaRiga ? 0 : 1), colonnaInizioAllDati + oreGiorno - 1]];
                        allDati.Style = "allDatiStyle";
                        //allDati.Name = GetName(info["SiglaEntita"], suffissoData, "ALLDATI" + allDatiIndice);
                        allDati.EntireColumn.ColumnWidth = _cell.Width.dato;
                        allDati.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
                        colonnaInizioAllDati += oreGiorno;
                        ultimaColonna = colonnaInizioAllDati + oreGiorno - 1;
                    });
                    allDatiIndice++;
                    rigaInizioGruppo = rigaAttiva + (info["SiglaTipologiaInformazione"].ToString() == "TITOLO2" ? 1 : 0);
                }
                if (primaRiga && primaRigaTitolo2)
                    rigaInizioGruppo++;

                rigaAttiva++;
            }

            rigaAttiva = _rigaAttiva;
            foreach (DataRowView info in informazioni)
            {
                if (!info["SiglaTipologiaInformazione"].Equals("TITOLO2"))
                {
                    string grassetto = "Bold:" + info["Grassetto"];
                    string formato = "NumberFormat:[" + info["Formato"] + "]";
                    string align = "Align:" + Enum.Parse(typeof(Excel.XlHAlign), info["Align"].ToString());

                    Excel.Range rigaInfo = _ws.Range[_ws.Cells[rigaAttiva, _colonnaInizio], _ws.Cells[rigaAttiva, ultimaColonna]];
                    Style.RangeStyle(rigaInfo, grassetto + ";" + formato + ";" + align);
                }
                rigaAttiva++;
            }
        }
        private void InsertInformazioniEntita(object siglaEntita, DataView informazioni)
        {
            int rigaAttiva = _rigaAttiva;
            int colonnaTitoloInfo = _colonnaInizio - VisParametro;

            bool titolo2 = false;
            foreach (DataRowView info in informazioni)
            {
                string bordoTop = "Top:" + (informazioni[0] == info || (info["InizioGruppo"].ToString() == "1" && !titolo2) ? "medium" : "thin");
                string bordoBottom = "Bottom:" + (informazioni[informazioni.Count - 1] == info ? "medium" : "thin");
                int backColor = (info["BackColor"] is DBNull ? 0 : (int)info["BackColor"]);
                backColor = backColor == 0 || backColor == 2 ? (info["Editabile"].ToString() == "1" ? 15 : 48) : backColor;
                titolo2 = false;

                //proprietà di stile comuni
                string style = "FontSize:" + info["FontSize"] + ";FontName:Verdana;BackColor:" + backColor + ";"
                    + "ForeColor:" + info["ForeColor"] + ";Visible:" + info["Visibile"] + ";";

                //personalizzazioni a seconda della tipologia di informazione
                if (info["SiglaTipologiaInformazione"].Equals("TITOLO2"))
                {
                    Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaTitoloInfo], _ws.Cells[rigaAttiva, colonnaTitoloInfo + _intervalloOre + 1]];
                    style += "Bold:" + info["Grassetto"] + ";Merge:true;Borders:[" + bordoTop + ",Bottom:thin,Right:medium]";
                    Style.RangeStyle(rng, style);
                    rng.Value = info["DesInformazione"].ToString();
                    titolo2 = true;
                }
                else
                {
                    Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaTitoloInfo], _ws.Cells[rigaAttiva, colonnaTitoloInfo + VisParametro - 1]];
                    style += "Borders:[insidev:thin,right:medium," + bordoTop + "," + bordoBottom + "]";
                    Style.RangeStyle(rng, style);

                    object[] valori = new object[VisParametro];
                    valori[0] = info["DesInformazione"];
                    valori[1] = info["DesInformazioneBreve"];

                    //TODO creare _struttura per COLONNA PARAMETRO                    
                    if (_struttura.visParametro)
                        valori[2] = "";

                    string nome = "";

                    if (!info["Selezione"].Equals("0"))
                    {
                        nome = (info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"]) + Simboli.UNION + "SEL" + info["Selezione"];
                    }

                    rng.Value = valori;
                    _ws.Cells[rigaAttiva, colonnaTitoloInfo + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    
                    if (info["Data0H24"].Equals("0"))
                        Style.RangeStyle(_ws.Cells[rigaAttiva, _colonnaInizio], "BackPattern:CrissCross");
                }
                rigaAttiva++;
            }
            rigaAttiva++;
        }
        private void CreaNomiCelle(DataView informazioni)
        {
            int rigaAttiva = _rigaAttiva;
            foreach (DataRowView info in informazioni)
            {
                int oraAttiva = _colonnaInizio;
                object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    bool isVisibleData0H24 = giorno == _dataInizio && _struttura.visData0H24;

                    if (isVisibleData0H24)
                    {
                        _nomiDefiniti.Add(DefinedNames.GetName(siglaEntita, info["SiglaInformazione"], "DATA0", "H24"), rigaAttiva, oraAttiva++);
                        oreGiorno--;
                    }

                    for (int i = 0; i < oreGiorno; i++)
                    {
                        _nomiDefiniti.Add(DefinedNames.GetName(siglaEntita, info["SiglaInformazione"], suffissoData, "H" + (i + 1)), rigaAttiva, oraAttiva++);
                    }
                });
                rigaAttiva++;
            }
        }
        private void InsertValoriCelle(DataView informazioni)
        {
            //carico tutti i dati reperibili durante la creazione del foglio
            int intervalloOre = GetOreIntervallo(_dataInizio, _dataFine);

            foreach (DataRowView info in informazioni)
            {
                object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                Tuple<int, int>[] riga = _nomiDefiniti.Get(DefinedNames.GetName(siglaEntita, info["SiglaInformazione"]), info["DATA0H24"].Equals("0"));

                if (info["ValoreDefault"] != DBNull.Value)
                {
                    Excel.Range rng = _ws.Range[_ws.Cells[riga[0].Item1, riga[0].Item2], _ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]];
                    rng.Value = info["ValoreDefault"];
                }
                else if (info["FormulaInCella"].Equals("1"))
                {
                    int deltaNeg;
                    int deltaPos;
                    string formula = "=" + PreparaFormula(info, "DATA0", "DATA1", 24, out deltaNeg, out deltaPos);

                    if (info["SiglaTipologiaInformazione"].Equals("OTTIMO"))
                    {
                        Tuple<int, int> cella = riga[0];

                        Excel.Range cellRng = _ws.Cells[cella.Item1, cella.Item2];
                        cellRng.Formula = "=SUM(" + Sheet.R1C1toA1(riga[1].Item1, riga[1].Item2) + ":" + Sheet.R1C1toA1(riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2) + ")";

                        Excel.Range rng = _ws.Range[_ws.Cells[riga[1].Item1, riga[1].Item2], _ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]];
                        rng.Formula = formula;
                    }
                    else
                    {
                        Excel.Range rng = _ws.Range[_ws.Cells[riga[0].Item1, riga[0].Item2 + deltaNeg], _ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2 - deltaPos]];
                        rng.Formula = formula;
                    }
                }
            }
        }
        //TODO controllare se funziona parametriH
        private void InsertParametri(DataView informazioni, DataView parametriD, DataView parametriH)
        {
            //carico tutti i dati reperibili durante la creazione del foglio
            int intervalloOre = GetOreIntervallo(_dataInizio, _dataFine);

            foreach (DataRowView info in informazioni)
            {
                int colonnaInizio = !_struttura.visData0H24 ? _colonnaInizio : _colonnaInizio + 1;
                object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                int rigaAttiva = _nomiDefiniti[DefinedNames.GetName(siglaEntita, info["SiglaInformazione"])][0].Item1;
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    oreGiorno = (giorno == _dataInizio && _struttura.visData0H24) ? oreGiorno - 1 : oreGiorno;

                    Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaInizio], _ws.Cells[rigaAttiva, colonnaInizio + oreGiorno - 1]];
                    parametriD.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaParametro = '" + info["SiglaTipologiaParametro"] + "' AND CONVERT(DataIV, System.Int32) <= " + giorno.ToString("yyyyMMdd") + " AND CONVERT(DataFV, System.Int32) >= " + giorno.ToString("yyyyMMdd");

                    if (parametriD.Count > 0)
                        rng.Value = parametriD[0]["Valore"];
                    else
                    {
                        parametriH.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaParametro = '" + info["SiglaTipologiaParametro"] + "' AND CONVERT(DataIV, System.Int32) <= " + giorno.ToString("yyyyMMdd") + " AND CONVERT(DataFV, System.Int32) >= " + giorno.ToString("yyyyMMdd");

                        parametriH.Sort = "Ora";

                        object[] values = parametriH.ToTable(false, "Valore").AsEnumerable().Select(r => r["Valore"]).ToArray();
                        rng.Value = values;
                    }

                    colonnaInizio += oreGiorno;
                });
                rigaAttiva++;
            }
        }
        private string PreparaFormula(DataRowView info, string suffissoDataPrec, string suffissoData, int oreDataPrec, out int deltaNeg, out int deltaPos)
        {
            if (info["Formula"] != DBNull.Value || info["Funzione"] != DBNull.Value)
            {
                string formula = info["Formula"] is DBNull ? info["Funzione"].ToString() : info["Formula"].ToString();

                string[] parametri = info["FormulaParametro"].ToString().Split(',');

                int tmpdeltaNeg = 0;
                int tmpdeltaPos = 0;

                foreach(string par in parametri) 
                {
                    if(Regex.IsMatch(par, @"\[[-+]?\d+\]"))
                    {
                        int deltaOre = int.Parse(par.Split('[')[1].Replace("]", ""));
                        if(deltaOre > 0)
                            tmpdeltaPos = Math.Max(tmpdeltaPos, deltaOre);
                        else
                            tmpdeltaNeg = Math.Min(tmpdeltaNeg, deltaOre);
                    }
                }
                
                deltaNeg = Math.Abs(tmpdeltaNeg);
                deltaPos = tmpdeltaPos;

                formula = Regex.Replace(formula, @"%P\d+(E\d+)?%", delegate(Match m)
                {
                    string[] parametroEntita = m.Value.Split('E');
                    int n = int.Parse(Regex.Match(parametroEntita[0], @"\d+").Value);

                    string nome = "";
                    if (parametroEntita.Length > 1)
                    {
                        int eRif = int.Parse(Regex.Match(parametroEntita[1], @"\d+").Value);
                        DataView categoriaEntita = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;
                        categoriaEntita.RowFilter = "Gerarchia = '" + info["SiglaEntita"] + "' AND Riferimento = " + eRif;
                        nome = DefinedNames.GetName(categoriaEntita[0]["SiglaEntita"], parametri[n - 1]);
                    }
                    else
                        nome = DefinedNames.GetName(info["SiglaEntita"], parametri[n - 1]);

                    if(Regex.IsMatch(nome, @"\[[-+]?\d+\]"))
                    {
                        int deltaOre = int.Parse(nome.Split('[')[1].Replace("]", ""));

                        if (suffissoData == "DATA1")
                        {//traslo in avanti la formula di |deltaNeg| - |deltaOre|
                            int ora = Math.Abs(tmpdeltaNeg) + deltaOre + (info["Data0H24"].Equals("1") ? 0 : 1);
                            nome += Simboli.UNION + (ora  == 0 ? DefinedNames.GetName("DATA0", "H24") : DefinedNames.GetName("DATA1", "H" + ora));
                        } 
                        else 
                        {
                            int ora = (deltaOre < 0 ? oreDataPrec : Math.Abs(tmpdeltaNeg)) + deltaOre + 1;
                            nome += Simboli.UNION + DefinedNames.GetName(deltaOre < 0 ? suffissoDataPrec : suffissoData, "H" + ora);
                        }
                        nome = Regex.Replace(nome, @"\[[-+]?\d+\]", "");
                    }
                    else
                    {
                        int ora = tmpdeltaNeg == 0 ? 1 : Math.Abs(tmpdeltaNeg) + (info["Data0H24"].Equals("1") ? 0 : 1);
                        nome += Simboli.UNION + DefinedNames.GetName(suffissoData, "H" + ora);
                    }

                    Tuple<int, int> coordinate = _nomiDefiniti[nome][0];
                        
                    return R1C1toA1(coordinate.Item1, coordinate.Item2);
                }, RegexOptions.IgnoreCase);
                return formula;
            }

            deltaNeg = 0;
            deltaPos = 0;
            
            return "";
        }
        private void CreaFormattazioneCondizionale(DataView informazioni, DataView formattazione)
        {
            int intervalloOre = GetOreIntervallo(_dataInizio, _dataFine);
            int colonnaInizio = !_struttura.visData0H24 ? _colonnaInizio : _colonnaInizio + 1;

            foreach (DataRowView info in informazioni)
            {
                object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                formattazione.RowFilter = (info["SiglaEntitaRif"] is DBNull ? "SiglaEntita" : "SiglaEntitaRif") + " = '" + siglaEntita + "' AND SiglaInformazione = '" + info["SiglaInformazione"] + "'";

                foreach (DataRowView format in formattazione)
                {
                    int rigaAttiva = _nomiDefiniti[DefinedNames.GetName(siglaEntita, info["SiglaInformazione"])][0].Item1;
                    Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaInizio], _ws.Cells[rigaAttiva, colonnaInizio + intervalloOre - 1]];

                    string[] valore = format["Valore"].ToString().Replace("\"", "").Split('|');
                    if (format["NomeCella"] != DBNull.Value)
                    {
                        Tuple<int, int> coordinate = _nomiDefiniti[DefinedNames.GetName(siglaEntita, format["NomeCella"], "DATA1", "H1")][0];
                        string address = R1C1toA1(coordinate.Item1, coordinate.Item2);

                        string formula = "";
                        switch ((int)format["Operatore"])
                        {
                            case 1:
                                formula = "=E(" + address + ">=" + valore[0] + ";" + address + "<=" + valore[1] + ")";
                                break;
                            case 3:
                                formula = "=" + address + "=" + valore[0];
                                break;
                            case 5:
                                formula = "=" + address + ">" + valore[0];
                                break;
                            case 6:
                                formula = "=" + address + "<" + valore[0];
                                break;
                        }
                        Excel.FormatCondition cond = rng.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: formula);

                        cond.Font.Color = format["ForeColor"];
                        cond.Font.Bold = format["Grassetto"].Equals("1");
                        if ((int)format["BackColor"] != 0)
                            cond.Interior.Color = format["BackColor"];
                        cond.Interior.Pattern = format["Pattern"];
                    }
                    else
                    {
                        string formula1;
                        string formula2 = "";
                        if ((int)format["Operatore"] == 1)
                        {
                            formula1 = valore[0];
                            formula2 = valore[1];
                        }
                        else
                        {
                            formula1 = valore[0];
                        }

                        Excel.FormatCondition cond = rng.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, format["Operatore"], formula1, formula2);

                        cond.Font.Color = format["ForeColor"];
                        cond.Font.Bold = format["Grassetto"].Equals("1");
                        if ((int)format["BackColor"] != 0)
                            cond.Interior.Color = format["BackColor"];

                        cond.Interior.Pattern = format["Pattern"];
                    }
                }
            }
        }

        #endregion

        //cancella i dati di tutti i giorni o del giorno specificato (attenzione ad usare in caso di cambio data perché il prefisso viene calcolato da data inizio config)
        private void CancellaDati(DateTime? giorno = null)
        {
            DataView dvCE = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;
            DataView dvEP = LocalDB.Tables[Tab.ENTITAPROPRIETA].DefaultView;

            dvCE.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "'"; // AND (Gerarchia = '' OR Gerarchia IS NULL )";

            _dataInizio = (DateTime)_config["DataInizio"];
            string suffissoData = giorno == null ? null : GetSuffissoData(_dataInizio, giorno.Value);

            foreach (DataRowView entita in dvCE)
            {
                DataView informazioni = LocalDB.Tables[Tab.ENTITAINFORMAZIONE].DefaultView;
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND FormulaInCella = '0' AND ValoreDefault IS NULL";
                
                foreach (DataRowView info in informazioni)
                {
                    var siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                    string nome = DefinedNames.GetName(siglaEntita, info["SiglaInformazione"], suffissoData);
                    //string nome = suffissoData == null ? DefinedNames.GetName(siglaEntita, info["SiglaInformazione"]) : DefinedNames.GetName(siglaEntita, info["SiglaInformazione"], suffissoData);
                    Tuple<int, int>[] riga = _nomiDefiniti[nome];

                    Excel.Range rng = _ws.Range[_ws.Cells[riga[0].Item1, riga[0].Item2], _ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]];
                    rng.Value = "";
                    rng.ClearComments();
                    Style.RangeStyle(rng, "BackColor:" + info["BackColor"] + ";ForeColor:" + info["ForeColor"]);
                }
            }
        }

        private void AggiornaDateTitoli()
        {
            DataView dvCE = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;
            DataView dvEP = LocalDB.Tables[Tab.ENTITAPROPRIETA].DefaultView;

            dvCE.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "' AND Gerarchia IS NULL";

            foreach (DataRowView entita in dvCE)
            {
                dvEP.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                if (dvEP.Count > 0)
                    _dataFine = _dataInizio.AddDays(double.Parse("" + dvEP[0]["Valore"]));
                else
                    _dataFine = _dataInizio.AddDays(_struttura.intervalloGiorni);

                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    Tuple<int, int>[] range = _nomiDefiniti.GetRange(DefinedNames.GetName(entita["SiglaEntita"], suffissoData));
                    _ws.Range[_ws.Cells[range[0].Item1, range[0].Item2], _ws.Cells[range[1].Item1, range[1].Item2]].Value = giorno;
                });
            }
        }

        public void CaricaParametri()
        {
            DataView dvEP = LocalDB.Tables[Tab.ENTITAPROPRIETA].DefaultView;
            DataView dvCE = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;

            DataView informazioni = LocalDB.Tables[Tab.ENTITAINFORMAZIONE].DefaultView;

            DataView parametriD = LocalDB.Tables[Tab.ENTITAPARAMETROD].DefaultView;
            DataView parametriH = LocalDB.Tables[Tab.ENTITAPARAMETROH].DefaultView;
            

            dvCE.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";
            _dataInizio = (DateTime)_config["DataInizio"];

            foreach (DataRowView entita in dvCE)
            {
                dvEP.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                if (dvEP.Count > 0)
                    _dataFine = _dataInizio.AddDays(double.Parse("" + dvEP[0]["Valore"]));
                else
                    _dataFine = _dataInizio.AddDays(_struttura.intervalloGiorni);

                informazioni.RowFilter += " AND SiglaTipologiaParametro IS NOT NULL";
                InsertParametri(informazioni, parametriD, parametriH);
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
            }
        }

        private void InsertGrafici()
        {
            DataView dvCE = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;
            DataView grafici = LocalDB.Tables[Tab.ENTITAGRAFICO].DefaultView;
            DataView graficiInfo = LocalDB.Tables[Tab.ENTITAGRAFICOINFORMAZIONE].DefaultView;

            dvCE.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "'";

            foreach (DataRowView entita in dvCE)
            {
                grafici.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
                graficiInfo.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

                int i = 1;
                foreach (DataRowView grafico in grafici)
                {
                    string nome = DefinedNames.GetName(grafico["SiglaEntita"], "GRAFICO" + i++);
                    Tuple<int, int>[] rangeGrafico = _nomiDefiniti.GetRange(nome);

                    var cella = _ws.Cells[rangeGrafico[0].Item1, rangeGrafico[0].Item2];

                    var rigaGrafico = _ws.Range[_ws.Cells[rangeGrafico[0].Item1, rangeGrafico[0].Item2], _ws.Cells[rangeGrafico[1].Item1, rangeGrafico[1].Item2]];
                    var chart = _ws.ChartObjects().Add(rigaGrafico.Left, rigaGrafico.Top + 1, rigaGrafico.Width, rigaGrafico.Height - 2).Chart;

                    chart.Parent.Name = nome;

                    chart.Axes(Excel.XlAxisType.xlCategory).TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
                    chart.Axes(Excel.XlAxisType.xlValue).HasMajorGridlines = false;
                    chart.Axes(Excel.XlAxisType.xlValue).HasMinorGridlines = false;
                    chart.Axes(Excel.XlAxisType.xlValue).MinorTickMark = Excel.XlTickMark.xlTickMarkOutside;
                    chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Name = "Verdana";
                    chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Size = 11;
                    chart.Axes(Excel.XlAxisType.xlValue).TickLabels.NumberFormat = "general";

                    chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;
                    chart.HasDataTable = false;
                    chart.DisplayBlanksAs = Excel.XlDisplayBlanksAs.xlNotPlotted;
                    chart.ChartGroups(1).GapWidth = 0;
                    chart.ChartGroups(1).Overlap = 100;
                    chart.ChartArea.Border.ColorIndex = 1;
                    chart.ChartArea.Border.Weight = 3;
                    chart.ChartArea.Border.LineStyle = 0;

                    chart.PlotArea.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                    foreach (DataRowView info in graficiInfo)
                    {
                        Tuple<int, int>[] rangeDati = _nomiDefiniti[DefinedNames.GetName(grafico["SiglaEntita"], info["SiglaInformazione"])];
                        var datiGrafico = _ws.Range[_ws.Cells[rangeDati[0].Item1, rangeDati[0].Item2], _ws.Cells[rangeDati[0].Item1, rangeDati[rangeDati.Length - 1].Item2]];
                        var serie = chart.SeriesCollection().Add(datiGrafico);
                        serie.Name = info["DesInformazione"].ToString();
                        serie.ChartType = (Excel.XlChartType)info["ChartType"];
                        serie.Interior.ColorIndex = info["InteriorColor"];
                        serie.Border.ColorIndex = info["BorderColor"];
                        serie.Border.Weight = info["BorderWeight"];
                        serie.Border.LineStyle = info["BorderLineStyle"];
                    }

                    AggiornaGrafici(chart, rigaGrafico);
                }
            }
        }

        private void AggiornaGrafici(Excel.Chart chart, Excel.Range rigaGrafico)
        {
            //resize dell'area del grafico per adattarla alle ore
            string max = chart.Axes(Excel.XlAxisType.xlValue).MaximumScale.ToString();
            string min = chart.Axes(Excel.XlAxisType.xlValue).MinimumScale.ToString();

            max = max.Length > min.Length ? max : min;

            Graphics grfx = Graphics.FromImage(new Bitmap(1, 1));
            grfx.PageUnit = GraphicsUnit.Point;
            SizeF sizeMax = grfx.MeasureString(max, new Font("Verdana", 11));

            chart.ChartArea.Left = rigaGrafico.Left - sizeMax.Width - 7;
            chart.ChartArea.Width = rigaGrafico.Width + sizeMax.Width + 4;
            chart.PlotArea.InsideLeft = 0;
            chart.PlotArea.Width = chart.ChartArea.Width + 3;
        }

        private void AggiornaGrafici()
        {
            DataView dvCE = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;
            DataView grafici = LocalDB.Tables[Tab.ENTITAGRAFICO].DefaultView;
            DataView graficiInfo = LocalDB.Tables[Tab.ENTITAGRAFICOINFORMAZIONE].DefaultView;

            dvCE.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "'";

            foreach (DataRowView entita in dvCE)
            {
                grafici.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
                graficiInfo.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

                int i = 1;
                foreach (DataRowView grafico in grafici)
                {
                    string nome = DefinedNames.GetName(grafico["SiglaEntita"], "GRAFICO" + i++);

                    Tuple<int, int>[] rangeGrafico = _nomiDefiniti.GetRange(nome);

                    var cella = _ws.Cells[rangeGrafico[0].Item1, rangeGrafico[0].Item2];
                    Excel.Range rigaGrafico = _ws.Range[_ws.Cells[rangeGrafico[0].Item1, rangeGrafico[0].Item2], _ws.Cells[rangeGrafico[1].Item1, rangeGrafico[1].Item2]];
                    var chart = _ws.ChartObjects(nome).Chart;

                    AggiornaGrafici(chart, rigaGrafico);
                }
            }
        }

        public void CaricaInformazioni()
        {
            DataView dvCE = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;
            DataView dvEP = LocalDB.Tables[Tab.ENTITAPROPRIETA].DefaultView;

            dvCE.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "'"; // AND (Gerarchia = '' OR Gerarchia IS NULL )";
            _dataInizio = (DateTime)_config["DataInizio"];

            //calcolo tutte le date e mantengo anche la data max
            DateTime dataFineMax = _dataInizio;
            Dictionary<object, DateTime> dateFineUP = new Dictionary<object, DateTime>();
            foreach (DataRowView entita in dvCE)
            {
                dvEP.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                if (dvEP.Count > 0)
                    dateFineUP.Add(entita["SiglaEntita"], _dataInizio.AddDays(double.Parse("" + dvEP[0]["Valore"])));
                else
                    dateFineUP.Add(entita["SiglaEntita"], _dataInizio.AddDays(_struttura.intervalloGiorni));

                dataFineMax = new DateTime(Math.Max(dataFineMax.Ticks, dateFineUP[entita["SiglaEntita"]].Ticks));
            }

            //Stopwatch watch = Stopwatch.StartNew();
            DataView datiApplicazione = DB.Select("spApplicazioneInformazione_test", "@SiglaCategoria=" + _config["SiglaCategoria"] + ";@SiglaEntita=ALL;@DateFrom=" + _dataInizio.ToString("yyyyMMdd") + ";@DateTo=" + dataFineMax.ToString("yyyyMMdd") + ";@All=1").DefaultView;

            DataView insertManuali = DB.Select("spApplicazioneInformazioneCommento_Test", "@SiglaCategoria=" + _config["SiglaCategoria"] + ";@SiglaEntita=ALL;@DateFrom=" + _dataInizio.ToString("yyyyMMdd") + ";@DateTo=" + dataFineMax.ToString("yyyyMMdd")).DefaultView;
            //watch.Stop();

            //dvCE.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";
            foreach (DataRowView entita in dvCE)
            {
                //watch = Stopwatch.StartNew();
                datiApplicazione.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND CONVERT(Data, System.Int32) <= " + dateFineUP[entita["SiglaEntita"]].ToString("yyyyMMdd");
                _dataFine = dateFineUP[entita["SiglaEntita"]];
                CaricaInformazioniEntita(datiApplicazione);
                //watch.Stop();
                //watch = Stopwatch.StartNew();
                insertManuali.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND CONVERT(SUBSTRING(Data, 1, 8), System.Int32) <= " + dateFineUP[entita["SiglaEntita"]].ToString("yyyyMMdd");
                CaricaCommentiEntita(insertManuali);
                //watch.Stop();
            }

        }
        #region Informazioni

        private void CaricaInformazioniEntita(DataView datiApplicazione)
        {
            foreach (DataRowView dato in datiApplicazione)
            {
                DateTime dataDato = DateTime.ParseExact(dato["Data"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                //sono nel caso DATA0H24
                if(dataDato < _dataInizio) 
                {
                    Tuple<int, int> cella = _nomiDefiniti[DefinedNames.GetName(dato["SiglaEntita"], dato["SiglaInformazione"], "DATA0", "H24")][0];
                    _ws.Cells[cella.Item1, cella.Item2].Value = dato["H24"];
                } 
                else 
                {
                    Tuple<int, int>[] riga = _nomiDefiniti[DefinedNames.GetName(dato["SiglaEntita"], dato["SiglaInformazione"], GetSuffissoData(_dataInizio, dataDato))];

                    List<object> o = new List<object>(dato.Row.ItemArray);
                    o.RemoveRange(0, 3);

                    _ws.Range[_ws.Cells[riga[0].Item1, riga[0].Item2], _ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Value = o.ToArray();
                }
            }
        }
        private void CaricaCommentiEntita(DataView insertManuali)
        {
            foreach (DataRowView commento in insertManuali)
            {
                DateTime giorno = DateTime.ParseExact(commento["Data"].ToString().Substring(0, 8), "yyyyMMdd", CultureInfo.InvariantCulture);
                Tuple<int, int> cella = _nomiDefiniti[DefinedNames.GetName(commento["SiglaEntita"], commento["SiglaInformazione"], GetSuffissoData(_dataInizio, giorno), GetSuffissoOra(commento["Data"]))][0];
                Excel.Range rng = _ws.Cells[cella.Item1, cella.Item2];
                rng.ClearComments();
                rng.AddComment("Valore inserito manualmente");
            }
        }

        #endregion

        public void CalcolaFormule(string siglaEntita = null, DateTime? dataAttiva = null, int ordineElaborazione = 0, bool escludiOrdine = false)
        {
            DataView dvCE = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;
            DataView dvEP = LocalDB.Tables[Tab.ENTITAPROPRIETA].DefaultView;
            DataView informazioni = LocalDB.Tables[Tab.ENTITAINFORMAZIONE].DefaultView;

            dvCE.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "' AND (Gerarchia = '' OR Gerarchia IS NULL )" + (siglaEntita == null ? "" : " AND SiglaEntita = '" + siglaEntita + "'");

            _dataInizio = (DateTime)_config["DataInizio"];
            DateTime giorno = dataAttiva ?? _dataInizio;

            foreach (DataRowView entita in dvCE)
            {
                siglaEntita = entita["SiglaEntita"].ToString();

                informazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND OrdineElaborazione <> 0 AND FormulaInCella = 0";
                if (ordineElaborazione != 0)
                {
                    informazioni.RowFilter += " AND OrdineElaborazione" + (escludiOrdine ? " <> " : " = ") + ordineElaborazione;
                }
                informazioni.Sort = "OrdineElaborazione";

                if (informazioni.Count > 0)
                {
                    if (dataAttiva == null)
                    {
                        dvEP.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                        if (dvEP.Count > 0)
                            _dataFine = _dataInizio.AddDays(double.Parse("" + dvEP[0]["Valore"]));
                        else
                            _dataFine = _dataInizio.AddDays(_struttura.intervalloGiorni);
                    }
                    else
                    {
                        _dataFine = giorno;
                    }

                    int intervalloOre = GetOreIntervallo(giorno, _dataFine);
                    int colonnaInizio = (_struttura.visData0H24 && giorno == _dataInizio) ? _colonnaInizio + 1 : _colonnaInizio;

                    string suffissoDataPrec = GetSuffissoData(_dataInizio, giorno.AddDays(-1));
                    string suffissoData = GetSuffissoData(_dataInizio, giorno);

                    foreach (DataRowView info in informazioni)
                    {
                        int rigaAttiva = _nomiDefiniti[DefinedNames.GetName(entita["SiglaEntita"], info["SiglaInformazione"])][0].Item1;

                        int deltaNeg;
                        int deltaPos;
                        string formula = "=" + PreparaFormula(info, suffissoDataPrec, suffissoData, GetOreGiorno(giorno.AddDays(-1)), out deltaNeg, out deltaPos);

                        Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaInizio - deltaNeg], _ws.Cells[rigaAttiva, colonnaInizio + intervalloOre - deltaPos - 1]];

                        rng.Formula = formula;

                        rigaAttiva++;
                    }
                }
                informazioni.Sort = "";
            }

        }

        #region Metodi Statici

        public static void Proteggi(bool proteggi)
        {
            foreach (Excel.Worksheet ws in WB.Sheets)
            {
                if (proteggi)
                    if(ws.Name == "Log")
                        ws.Protect(Simboli.pwd, AllowSorting: true, AllowFiltering: true);
                    else
                        ws.Protect(Simboli.pwd);
                else
                    ws.Unprotect(Simboli.pwd);
            }
        }

        public static void AbilitaModifica(bool abilita)
        {
            DataView categorie = LocalDB.Tables[Tab.CATEGORIA].DefaultView;
            categorie.RowFilter = "Operativa = '1'";
            DataView entita = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;
            
            DataView informazioni = LocalDB.Tables[Tab.ENTITAINFORMAZIONE].DefaultView;

            foreach (DataRowView categoria in categorie)
            {
                DefinedNames nomiDefiniti = new DefinedNames(categoria["DesCategoria"].ToString());
                Excel.Worksheet ws = CommonFunctions.WB.Sheets[categoria["DesCategoria"].ToString()];

                ws.Unprotect(Simboli.pwd);

                entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "'";
                foreach (DataRowView e in entita)
                {
                    informazioni.RowFilter = "SiglaEntita = '" + e["SiglaEntita"] + "' AND Editabile = '1'";
                    foreach (DataRowView info in informazioni)
                    {
                        object siglaEntita = info["SiglaEntitaRif"] is DBNull ? e["SiglaEntita"] : info["SiglaEntitaRif"];
                        Tuple<int, int>[] riga = nomiDefiniti[DefinedNames.GetName(siglaEntita, info["SiglaInformazione"])];

                        ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Locked = !abilita;
                    }
                }
                ws.Protect(Simboli.pwd);
            }
        }

        public static string R1C1toA1(int riga, int colonna)
        {
            string output = "";
            while (colonna > 0)
            {
                int lettera = (colonna % 26 == 0 ? 26 : colonna % 26);
                output = char.ConvertFromUtf32(lettera + 64) + output;
                colonna = colonna / 26;
            }
            output += riga;
            return output;
        }

        #endregion


        public void Dispose()
        {
            if (!_disposed)
            {
                GC.SuppressFinalize(this);
                _disposed = true;
            }
        }
    }
}
