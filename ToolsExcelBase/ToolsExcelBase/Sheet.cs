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

namespace Iren.FrontOffice.Base
{
    public class Sheet<T> : CommonFunctions, IDisposable
    {
        #region Variabili

        Worksheet _ws;
        Dictionary<string, object> _config;
        DateTime _dataInizio;
        DateTime _dataFine;
        int _colonnaInizio;
        int _intervalloOre;
        int _rigaAttiva;
        string _nomeFoglio;
        bool _disposed = false;
        DataTable _nomiDefiniti;
        
        #endregion

        #region Costruttori

        public Sheet(T categoria)
        {
            Type t = categoria.GetType();
            PropertyInfo p = t.GetProperty("Base");
            _ws = (Worksheet)p.GetValue(categoria, null);

            FieldInfo f = t.GetField("config");
            _config = (Dictionary<string, object>)f.GetValue(categoria);

            //dimensionamento celle in base ai parametri del DB
            DataView paramApplicazione = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.APPLICAZIONE].DefaultView;

            Cell.Width.empty = double.Parse(paramApplicazione[0]["ColVuotaWidth"].ToString());
            Cell.Width.dato = double.Parse(paramApplicazione[0]["ColDatoWidth"].ToString());
            Cell.Width.entita = double.Parse(paramApplicazione[0]["ColEntitaWidth"].ToString());
            Cell.Width.informazione = double.Parse(paramApplicazione[0]["ColInformazioneWidth"].ToString());
            Cell.Width.unitaMisura = double.Parse(paramApplicazione[0]["ColUMWidth"].ToString());
            Cell.Width.parametro = double.Parse(paramApplicazione[0]["ColParametroWidth"].ToString());
            Cell.Height.normal = double.Parse(paramApplicazione[0]["RowHeight"].ToString());
            Cell.Height.empty = double.Parse(paramApplicazione[0]["RowVuotaHeight"].ToString());
            Struttura.rigaBlock = (int)paramApplicazione[0]["RowBlocco"];
            Struttura.rigaGoto = (int)paramApplicazione[0]["RowGoto"];
            Struttura.intervalloGiorni = (int)paramApplicazione[0]["IntervalloGiorni"];
            Struttura.visData0H24 = paramApplicazione[0]["VisData0H24"].ToString() == "1";
            Struttura.visParametro = paramApplicazione[0]["VisParametro"].ToString() == "1";
            Struttura.colBlock = (int)paramApplicazione[0]["ColBlocco"] + (Struttura.visParametro ? 1 : 0);

            Style.StdStyles(CommonFunctions.ThisWorkBook);

            _nomiDefiniti = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.NOMIDEFINITI];
        }
        ~Sheet()
        {
            Dispose();
        }

        #endregion


        private delegate void CicloGiorni(int oreGiorno, string suffissoData, DateTime giorno);
        private void EseguiCicloGiorni(CicloGiorni callback)
        {
            for (DateTime giorno = _dataInizio; giorno <= _dataFine; giorno = giorno.AddDays(1))
            {
                int oreGiorno = GetOreGiorno(giorno);
                string suffissoData = GetSuffissoData(_dataInizio, giorno);

                if (giorno == _dataInizio && Struttura.visData0H24)
                {
                    oreGiorno++;
                }
                
                callback(oreGiorno, suffissoData, giorno);
            }
        }

        private void DefineNewName(string nome, Tuple<int, int> cella1, Tuple<int, int> cella2 = null)
        {
            DataRow r = _nomiDefiniti.NewRow();
            r["Nome"] = nome;
            r["Cella1"] = cella1;
            r["Cella2"] = cella2 ?? cella1;

//TODO controllare se nome esiste già

            _nomiDefiniti.Rows.Add(r);
        }

        private void Clear()
        {
            int dataOreTot = GetOreIntervallo(_dataInizio, _dataInizio.AddDays(Struttura.intervalloGiorni)) + (Struttura.visData0H24 ? 1 : 0) + (Struttura.visParametro ? 1 : 0);

            _ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            _ws.UsedRange.EntireColumn.Delete();
            _ws.UsedRange.FormatConditions.Delete();
            _ws.UsedRange.EntireRow.Hidden = false;
            _ws.UsedRange.Font.Size = 10;
            _ws.UsedRange.Font.Name = "Verdana";
            _ws.UsedRange.RowHeight = Cell.Height.normal;

            _ws.Rows["1:" + (Struttura.rigaBlock - 1)].RowHeight = Cell.Height.empty;
            _ws.Rows[Struttura.rigaGoto].RowHeight = Cell.Height.normal;

            _ws.Columns[1].ColumnWidth = Cell.Width.empty;
            _ws.Columns[2].ColumnWidth = Cell.Width.entita;

            _ws.Activate();
            _ws.Application.ActiveWindow.FreezePanes = false;
            _ws.Cells[Struttura.rigaBlock, Struttura.colBlock].Select();
            _ws.Application.ActiveWindow.ScrollColumn = 1;
            _ws.Application.ActiveWindow.ScrollRow = 1;
            _ws.Application.ActiveWindow.FreezePanes = true;

            string gotoBarRangeName = _config["SiglaCategoria"] + Simboli.UNION + "GOTO_BAR";
            Excel.Range rng = _ws.Range[_ws.Cells[2, 2], _ws.Cells[Struttura.rigaBlock - 2, Struttura.colBlock + dataOreTot - 1]];
            rng.Style = "gotoBarStyle";
            rng.BorderAround2(Weight: Excel.XlBorderWeight.xlMedium, Color: 1);

            int infoCols = Struttura.visParametro ? 3 : 2;
            
            for(int i = infoCols; i > 0; i--) 
            {
                if (i % infoCols == 0)
                    _ws.Columns[Struttura.colBlock - i].ColumnWidth = Cell.Width.informazione;
                else if (i % (infoCols - 1) == 0)
                    _ws.Columns[Struttura.colBlock - i].ColumnWidth = Cell.Width.unitaMisura;
                else if (i % (infoCols - 2) == 0)
                    _ws.Columns[Struttura.colBlock - i].ColumnWidth = Cell.Width.parametro;
            }
        }

        public void LoadStructure()
        {
            DataView dvC = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.CATEGORIA].DefaultView;
            DataView dvCE = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.CATEGORIAENTITA].DefaultView;
            DataView dvEP = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAPROPRIETA].DefaultView;

            dvC.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "'";
            _nomeFoglio = dvC[0]["DesCategoria"].ToString();
            dvCE.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";
            _dataInizio = (DateTime)_config["DataInizio"];

            Clear();
            InitBarraNavigazione(dvCE);

            //int intervalloOre;
            _rigaAttiva = Struttura.rigaBlock + 1;

            foreach (DataRowView entita in dvCE)
            {
                string siglaEntita = ""+entita["SiglaEntita"];
                dvEP.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta LIKE '%GIORNI_STRUTTURA'";
                if (dvEP.Count > 0)
                    _dataFine = _dataInizio.AddDays(double.Parse("" + dvEP[0]["Valore"]));
                else
                    _dataFine = _dataInizio.AddDays(Struttura.intervalloGiorni);

                InitBloccoEntita(entita);                

            }
        }

        private void InitBarraNavigazione(DataView entita)
        {
            object[] descrizioni = new object[entita.Count];
            int i = -1;
            foreach (DataRowView e in entita)
            {
                descrizioni[++i] = e["DesEntitaBreve"];
                //_ws.Cells[Struttura.rigaGoto, Struttura.colBlock + i].Name = e["siglaEntita"] + Simboli.UNION + "GOTO";
                DefineNewName(e["siglaEntita"] + Simboli.UNION + "GOTO", Tuple.Create(Struttura.rigaGoto, Struttura.colBlock + i));
            }

            Excel.Range rng = _ws.Range[_ws.Cells[Struttura.rigaGoto, Struttura.colBlock],
                _ws.Cells[Struttura.rigaGoto, Struttura.colBlock + i]];
            //string gotoMenuRangeName = _config["SiglaCategoria"] + Simboli.UNION + "GOTO_MENU";

            rng.Value = descrizioni;
            rng.Style = "navBarStyle";
        }

        private void InitBloccoEntita(DataRowView entita)
        {
            _rigaAttiva++;
            DataView grafici = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAGRAFICO].DefaultView;
            grafici.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

            DataView informazioni = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAINFORMAZIONE].DefaultView;
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

            _colonnaInizio = Struttura.colBlock;
            _intervalloOre = CommonFunctions.GetOreIntervallo(_dataInizio, _dataFine) + (Struttura.visData0H24 ? 1 : 0) + (Struttura.visParametro ? 1 : 0);

            //titolo + data
            InsertTitoloEntita(entita);
            //grafici
            InsertGraficiEntita(grafici);
            //ore
            InsertOre(entita["SiglaEntita"]);
            //titolo verticale
            InsertTitoloVerticale(entita["SiglaEntita"], entita["DesEntitaBreve"], informazioni.Count);
            //informazioni
            InsertInformazioniEntita(entita["SiglaEntita"], informazioni);
            //ore
            InsertNomiValori(informazioni);
            //formattazione condizionale
            CreateFormattazioneCondizionale(informazioni);
        }
        #region Blocco entità
        
        private void InsertTitoloEntita(DataRowView entita)
        {
            int colonnaInizio = _colonnaInizio;            

            EseguiCicloGiorni((oreGiorno, suffissoData, giorno) =>
               {
                   string rangeTitolo = entita["SiglaEntita"] + Simboli.UNION + "T" + Simboli.UNION + suffissoData;

                   Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, colonnaInizio],
                           _ws.Cells[_rigaAttiva, colonnaInizio + oreGiorno - 1]];

                   DefineNewName(rangeTitolo, Tuple.Create(_rigaAttiva, colonnaInizio), Tuple.Create(_rigaAttiva, colonnaInizio + oreGiorno - 1));

                   rng.Merge();
                   rng.Style = "titleBarStyle";
                   rng.Value = entita["DesEntita"].ToString().ToUpperInvariant();
                   rng.RowHeight = 25;

                   rng = _ws.Range[_ws.Cells[_rigaAttiva + 1, colonnaInizio],
                           _ws.Cells[_rigaAttiva + 1, colonnaInizio + oreGiorno - 1]];

                   rng.Merge();
                   rng.Style = "dateBarStyle";
                   rng.Value = giorno.ToString("MM/dd/yyyy");
                   rng.RowHeight = 20;

                   colonnaInizio += oreGiorno;
               });
            _rigaAttiva++;
        }
        private void InsertGraficiEntita(DataView grafici)
        {            
            int i = 1;
            foreach (DataRowView grafico in grafici)
            {
                string graficoRange = grafico["SiglaEntita"] + Simboli.UNION + "GRAFICO" + (grafici.Count > 1 ? ""+i++ : "");

                Excel.Range rng = _ws.Range[_ws.Cells[++_rigaAttiva, _colonnaInizio],
                    _ws.Cells[_rigaAttiva, _colonnaInizio + _intervalloOre - 1]];

                DefineNewName(graficoRange, Tuple.Create(++_rigaAttiva, _colonnaInizio), Tuple.Create(_rigaAttiva, _colonnaInizio + _intervalloOre - 1));

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

            EseguiCicloGiorni(delegate(int oreGiorno, string suffissoData, DateTime giorno) 
                {
                    Style.RangeStyle(_ws.Range[_ws.Cells[_rigaAttiva, colonnaInizio], _ws.Cells[_rigaAttiva, colonnaInizio + oreGiorno - 1]], style);
                    for (int i = colonnaInizio; i < colonnaInizio + oreGiorno; i++)
                    {
                        int val = i - colonnaInizio + 1;
                        if (giorno == _dataInizio && Struttura.visData0H24)
                            val = i == colonnaInizio ? 24 : i - colonnaInizio;
                        _ws.Cells[_rigaAttiva, i].Value = val;
                    }
                    colonnaInizio += oreGiorno;
                });
            _rigaAttiva++;
        }
        private void InsertTitoloVerticale(object siglaEntita, object siglaEntitaBreve, int numInformazioni)
        {
            int colonnaTitoloVert = _colonnaInizio - (Struttura.visParametro ? 4 : 3);
            Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, colonnaTitoloVert], _ws.Cells[_rigaAttiva + numInformazioni - 1, colonnaTitoloVert]];
            rng.Style = "titoloVertStyle";
            rng.Merge();
            rng.Orientation = numInformazioni == 1 ? Excel.XlOrientation.xlHorizontal : Excel.XlOrientation.xlVertical;
            rng.Font.Size = numInformazioni == 1 ? 6 : 9;
            rng.Value = siglaEntitaBreve;
        }
        private void InsertInformazioniEntita(object siglaEntita, DataView informazioni)
        {
            //creo e formatto il range ALLDATI
            int colonnaInizio = _colonnaInizio;
            int rigaAttiva = _rigaAttiva;

//TODO migliorare allData ranges!!!!! non serve definire nomi per questi

            //EseguiCicloGiorni(delegate(int oreGiorno, string suffissoData, DateTime giorno)
            //    {
            //        Excel.Range allDati = _ws.Range[_ws.Cells[rigaAttiva, colonnaInizio], _ws.Cells[rigaAttiva + informazioni.Count - 1, colonnaInizio + oreGiorno - 1]];
            //        allDati.Style = "allDatiStyle";
            //        allDati.Name = siglaEntita + Simboli.UNION + "ALLDATI" + Simboli.UNION + suffissoData;
            //        allDati.EntireColumn.ColumnWidth = Cell.Width.dato;
            //        allDati.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            //        colonnaInizio += oreGiorno;
            //    });

            int colonnaTitoloInfo = _colonnaInizio - (Struttura.visParametro ? 3 : 2);
            foreach (DataRowView info in informazioni)
            {
                string siglaEntitaInfo = (info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"]).ToString();
                string siglaInfo = info["SiglaInformazione"].ToString();
                string bordoTop = "Top:" + (informazioni[0] == info ||  info["InizioGruppo"].ToString() == "1" ? "medium" : "thin");
                string bordoBottom = "Bottom:" + (informazioni[informazioni.Count - 1] == info ? "medium" : "thin");
                bool grassetto = info["Grassetto"].ToString() == "1";
                int backColor = (info["BackColor"] is DBNull ? 0 : (int)info["BackColor"]);
                backColor = backColor == 0 || backColor == 2 ? (info["Editabile"].ToString() == "1" ? 15 : 48) : backColor;

                //proprietà di stile comuni
                string style = "FontSize=" + info["FontSize"] + ";BackColor=" + backColor + ";"
                    + "ForeColor=" + info["ForeColor"] + ";Visible=" + info["Visibile"] + ";";

                //personalizzazioni a seconda della tipologia di informazione
                if (info["SiglaTipologiaInformazione"].ToString() == "TITOLO2")
                {
                    Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaTitoloInfo], _ws.Cells[rigaAttiva, colonnaTitoloInfo + _intervalloOre + 1]];
                    style += "Bold:" + (grassetto) + ";Merge:true;Borders:[" + bordoTop + "," + bordoBottom + ", Right:medium]";
                    Style.RangeStyle(rng, style);
                    rng.Value = info["DesInformazione"].ToString();
                }
                else
                {
                    Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaTitoloInfo], _ws.Cells[rigaAttiva, colonnaTitoloInfo + (Struttura.visParametro ? 2 : 1)]];                    
                    style += "Borders:[insidev:thin,right:medium," + bordoTop + "," + bordoBottom + "]";
                    Style.RangeStyle(rng, style);

                    object[] valori = new object[(Struttura.visParametro ? 3 : 2)];
                    valori[0] = info["DesInformazione"];
                    valori[1] = info["DesInformazioneBreve"];
                    
//TODO creare struttura per COLONNA PARAMETRO                    
                    if (Struttura.visParametro) 
                        valori[2] = "";
                    
                    string nome = "";

                    if (info["Selezione"].ToString() != "0") 
                    {
                        nome = siglaEntitaInfo + Simboli.UNION + "SEL" + info["Selezione"];
                    }

                    //scrivo i valori nelle celle
                    rng.Value = valori;
                    //scrivo i nomi dove necessario
                    //_ws.Cells[rigaAttiva, colonnaTitoloInfo].Name = siglaEntitaInfo + Simboli.UNION + info["SiglaInformazione"];
                    //if(nome != "")
                    //    _ws.Cells[rigaAttiva, colonnaTitoloInfo + 1].Name = nome;
                    //cambio impostazioni per la seconda cella
                    _ws.Cells[rigaAttiva, colonnaTitoloInfo + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   
                }
                rigaAttiva++;
            }
            rigaAttiva++;
        }
        private void InsertNomiValori(DataView informazioni)
        {
            int rigaAttiva = _rigaAttiva;

            //TODO Rimettere a posto il ciclo formattando e riempiendo i valori a blocchi e non per cella
            foreach (DataRowView info in informazioni)
            {
                string bordoTop = "Top:" + (informazioni[0] == info || info["InizioGruppo"].ToString() == "1" ? "medium" : "thin");
                string bordoBottom = (informazioni[informazioni.Count - 1] == info ? "Bottom:medium" : "");
                bool grassetto = info["Grassetto"].ToString() == "1";
                int backColor = (info["BackColor"] is DBNull ? 0 : (int)info["BackColor"]);
                backColor = backColor == 0 || backColor == 2 ? (info["Editabile"].ToString() == "1" ? 15 : 48) : backColor;

                string formula = PreparaFormula(info);

                string style = "FontSize=" + info["FontSize"] + ";BackColor=" + backColor + ";"
                    + "ForeColor=" + info["ForeColor"] + ";Visible=" + info["Visibile"] + ";";

                int colonnaAttiva = _colonnaInizio - 1;

            }
            
            
            
            
            
            
            
            
            
            foreach (DataRowView info in informazioni)
            {
                string siglaEntitaInfo = (info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"]).ToString();
                string bordoTop = "Top:" + (informazioni[0] == info || info["InizioGruppo"].ToString() == "1" ? "medium" : "thin");
                string bordoBottom = (informazioni[informazioni.Count - 1] == info ? "Bottom:medium" : "");
                bool grassetto = info["Grassetto"].ToString() == "1";
                int backColor = (info["BackColor"] is DBNull ? 0 : (int)info["BackColor"]);
                backColor = backColor == 0 || backColor == 2 ? (info["Editabile"].ToString() == "1" ? 15 : 48) : backColor;

                string formula = PreparaFormula(info);

                //proprietà di stile comuni
                string style = "FontSize=" + info["FontSize"] + ";BackColor=" + backColor + ";"
                    + "ForeColor=" + info["ForeColor"] + ";Visible=" + info["Visibile"] + ";";

                //scrivo i nomi per le ore e personalizzo lo stile dove necessario
                int colonnaAttiva = _colonnaInizio - 1;
                EseguiCicloGiorni(delegate(int oreGiorno, string suffissoData, DateTime giorno)
                {
                    bool isVisibleData0H24 = giorno == _dataInizio && Struttura.visData0H24;
                    
                    if (isVisibleData0H24)
                    {
                        _ws.Cells[rigaAttiva, ++colonnaAttiva].Name = (info["SiglaEntita"] + Simboli.UNION + info["SiglaInformazione"] + Simboli.UNION + "DATA0" + Simboli.UNION + "H24");
                        style = "ForeColor:" + info["ForeColor"] + ";BackColor" + backColor + ";Bold:" + grassetto + ";NumberFormat:" + info["Formato"]
                            + "Borders:[InsideV:thin," + bordoTop + "," + bordoBottom + "]";
                        Style.RangeStyle(_ws.Cells[rigaAttiva, colonnaAttiva], style);

                        if (info["SiglaInformazione"].ToString() == "VOL_INVASO")
                        {
                            _ws.Cells[rigaAttiva, colonnaAttiva].Value = 160;
                        }
                        else if (info["SiglaInformazione"].ToString() == "TEMP_PROG5")
                        {
                            _ws.Cells[rigaAttiva, colonnaAttiva].Value = "=SUM($E$25:$FP$25)";
                        }
                    }
                    int oreRealiGiorno = oreGiorno - (isVisibleData0H24 ? 1 : 0);
                    for (int i = 1; i <= oreRealiGiorno; i++)
                    {
                        int oraAttiva = colonnaAttiva + i;

                        _ws.Cells[rigaAttiva, oraAttiva].Name = (info["SiglaEntita"] + Simboli.UNION + info["SiglaInformazione"] + Simboli.UNION + suffissoData + Simboli.UNION + "H" + (i));
                        style = "ForeColor:" + info["ForeColor"] + ";BackColor" + backColor + ";Bold:" + grassetto + ";NumberFormat:" + info["Formato"]
                            + "Borders:[InsideV:thin," + bordoTop + "," + bordoBottom + (i == oreRealiGiorno ? ",Right:medium" : "") + "]";
                        Style.RangeStyle(_ws.Cells[rigaAttiva, oraAttiva], style);

                        if (!(info["ValoreDefault"] is DBNull))
                        {
                            _ws.Cells[rigaAttiva, oraAttiva].FormulaR1C1 = double.Parse(info["ValoreDefault"].ToString().Replace('.', ','));
                        }
                        else
                        {
                            if (info["FormulaInCella"].Equals("1"))
                            {
                                string formulaFinale  = "=";
                                if (!(info["Formula"] is DBNull))
                                    formulaFinale += formula.Replace("%DATA%", suffissoData + Simboli.UNION + "H" + i);
                                else
                                {
                                    formulaFinale += formula.Replace("%DATA%", suffissoData + Simboli.UNION + "H" + i);
                                    if (siglaEntitaInfo == "UP_BUS")
                                    {
                                        string dataPrec = (i == 1 ? "DATA0" + Simboli.UNION + "H24" : suffissoData + Simboli.UNION + "H" + i);
                                        formulaFinale += formulaFinale.Replace("%DATA-1%", dataPrec);
                                    }
                                    _ws.Cells[rigaAttiva, oraAttiva].FormulaR1C1 = formulaFinale;
                                }
                            }
                        }
                    }
                    colonnaAttiva += oreRealiGiorno;
                });
                rigaAttiva++;
            }
            rigaAttiva++;
        }
        private string PreparaFormula(DataRowView info)
        {
            if (info["FormulaInCella"].Equals("1"))
            {
                string formula = info["Formula"].ToString();
                if (formula == "")
                    formula = info["Funzione"].ToString().Replace("%SHEET%", _nomeFoglio).Replace("%ENTITA%", info["SiglaEntita"].ToString());
                else
                {
                    string[] parametri = info["Parametro"].ToString().Split(',');
                    formula = Regex.Replace(formula, @"%P\d+%", delegate(Match m)
                        {
                            int n = int.Parse(Regex.Match(m.Value, @"\d+").Value);
                            string s = info["SiglaEntita"] + Simboli.UNION + parametri[n - 1].Replace("[-1]","");
                            return s.ToUpperInvariant() + Simboli.UNION + (parametri[n - 1].EndsWith("[-1]") ? "%DATA-1%" : "%DATA%");
                        }, RegexOptions.IgnoreCase);
                }
                return formula;
            }
            return "";
        }

        private void CreateFormattazioneCondizionale(DataView informazioni)
        {
            foreach (DataRowView info in informazioni)
            {
                _rigaAttiva++;
            }
            _rigaAttiva++;
        }

        #endregion


        #region Classi per la struttura delle pagina

        internal class Cell
        {
            public class Width
            {
                public static double empty = 1,
                dato = 8.8,
                entita = 3,
                informazione = 28,
                unitaMisura = 6,
                parametro = 8.8,
                riepilogo = 9;
            }

            public class Height
            {
                public static double normal = 15,
                empty = 5;
            }
        }

        internal class Struttura
        {
            public static int colBlock = 5,
                rigaBlock = 6,
                rigaGoto = 3,
                intervalloGiorni = 0;
            public static bool visData0H24 = false,
                visParametro = false;
        }

        internal class Simboli
        {
            public const string UNION = ".";
        }

        #endregion

        public void Dispose()
        {
            if (!_disposed)
            {
                _ws.Dispose();
                GC.SuppressFinalize(this);
                _disposed = true;
            }
        }
    }
}