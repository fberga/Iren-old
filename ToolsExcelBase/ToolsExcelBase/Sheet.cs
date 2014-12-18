﻿using System;
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
        //DataTable _nomiDefiniti;
        DefinedNames _nomiDefiniti;
        
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
            Struttura.intervalloGiorni = 2;// (int)paramApplicazione[0]["IntervalloGiorni"];
            Struttura.visData0H24 = paramApplicazione[0]["VisData0H24"].ToString() == "1";
            Struttura.visParametro = paramApplicazione[0]["VisParametro"].ToString() == "1";
            Struttura.colBlock = (int)paramApplicazione[0]["ColBlocco"] + (Struttura.visParametro ? 1 : 0);

            Style.StdStyles(CommonFunctions.ThisWorkBook);

            //_nomiDefiniti = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.NOMIDEFINITI];
            _nomiDefiniti = new DefinedNames();
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
                return Struttura.visParametro ? 3 : 2;
            }
        }

        #endregion

        private void CicloGiorni(Func<int, string, DateTime, bool> callback)
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

//        private void DefineNewName(string nome, Tuple<int, int> cella1, Tuple<int, int> cella2 = null)
//        {
//            DataRow r = _nomiDefiniti.NewRow();
//            cella2 = cella2 ?? cella1;
//            r["Nome"] = nome;
//            r["R1"] = cella1.Item1;
//            r["C1"] = cella1.Item2;
//            r["R2"] = cella2.Item1;
//            r["C2"] = cella2.Item2;

////TODO controllare se nome esiste già

//            _nomiDefiniti.Rows.Add(r);
//        }

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

            int infoCols = Struttura.colBlock - VisParametro;

            _ws.Columns[Struttura.colBlock].ColumnWidth = Cell.Width.informazione;
            _ws.Columns[Struttura.colBlock + 1].ColumnWidth = Cell.Width.unitaMisura;
            if(Struttura.visParametro)
                _ws.Columns[Struttura.colBlock + 2].ColumnWidth = Cell.Width.parametro;
        }

        public void LoadStructure()
        {
            DataView dvEP = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAPROPRIETA].DefaultView;
            DataView dvC = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.CATEGORIA].DefaultView;
            DataView dvCE = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.CATEGORIAENTITA].DefaultView;

            dvC.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "'";
            _nomeFoglio = dvC[0]["DesCategoria"].ToString();
            dvCE.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";
            _dataInizio = (DateTime)_config["DataInizio"];

            Clear();
            InitBarraNavigazione(dvCE);

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
                _nomiDefiniti.Add(e["siglaEntita"] + Simboli.UNION + "GOTO", Tuple.Create(Struttura.rigaGoto, Struttura.colBlock + i));
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
            //formatta AllDati
            FormattaAllDati(informazioni);
            //informazioni
            InsertInformazioniEntita(entita["SiglaEntita"], informazioni);
            //ore
            CreateNomiCelle(informazioni);
            InsertValoriCelle(informazioni);
            //InsertNomiValori(informazioni);
            //formattazione condizionale
            CreateFormattazioneCondizionale(informazioni);
        }
        #region Blocco entità
        
        private void InsertTitoloEntita(DataRowView entita)
        {
            int colonnaInizio = _colonnaInizio;            

            CicloGiorni((oreGiorno, suffissoData, giorno) =>
               {
                   string rangeTitolo = entita["SiglaEntita"] + Simboli.UNION + "T" + Simboli.UNION + suffissoData;

                   Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, colonnaInizio],
                           _ws.Cells[_rigaAttiva, colonnaInizio + oreGiorno - 1]];

                   _nomiDefiniti.Add(rangeTitolo, Tuple.Create(_rigaAttiva, colonnaInizio), Tuple.Create(_rigaAttiva, colonnaInizio + oreGiorno - 1));

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
                   
                   return true;
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
                        if (giorno == _dataInizio && Struttura.visData0H24)
                            val = i == colonnaInizio ? 24 : i - colonnaInizio;
                        _ws.Cells[_rigaAttiva, i].Value = val;
                    }
                    colonnaInizio += oreGiorno;

                    return true;
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
            int colonnaTitoloInfo = _colonnaInizio - VisParametro;
            int allDatiIndice = 1;

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
                            Style.RangeStyle(allDati, "Style:allDatiStyle;Bold:" + info["Grassetto"] + ";");
                            allDati.Name = info["SiglaEntita"] + Simboli.UNION + suffissoData + Simboli.UNION + "ALLDATI" + allDatiIndice;
                            allDati.EntireColumn.ColumnWidth = Cell.Width.dato;
                            allDati.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
                            colonnaInizioAllDati += oreGiorno;

                            return true;
                        });
                    allDatiIndice++;
                    rigaInizioGruppo = rigaAttiva + (info["SiglaTipologiaInformazione"].ToString() == "TITOLO2" ? 1 : 0);
                }
                rigaAttiva++;
            }
        }
        
        private void InsertInformazioniEntita(object siglaEntita, DataView informazioni)
        {
            int rigaAttiva = _rigaAttiva;
            int colonnaTitoloInfo = _colonnaInizio - VisParametro;
            
            foreach (DataRowView info in informazioni)
            {
                string bordoTop = "Top:" + (informazioni[0] == info || info["InizioGruppo"].ToString() == "1" ? "medium" : "thin");
                string bordoBottom = "Bottom:" + (informazioni[informazioni.Count - 1] == info ? "medium" : "thin");
                int backColor = (info["BackColor"] is DBNull ? 0 : (int)info["BackColor"]);
                backColor = backColor == 0 || backColor == 2 ? (info["Editabile"].ToString() == "1" ? 15 : 48) : backColor;

                //proprietà di stile comuni
                string style = "FontSize:" + info["FontSize"] + ";BackColor:" + backColor + ";"
                    + "ForeColor:" + info["ForeColor"] + ";Visible:" + info["Visibile"] + ";";

                //personalizzazioni a seconda della tipologia di informazione
                if (info["SiglaTipologiaInformazione"].ToString() == "TITOLO2")
                {
                    Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaTitoloInfo], _ws.Cells[rigaAttiva, colonnaTitoloInfo + _intervalloOre + 1]];
                    style += "Bold:" + info["Grassetto"] + ";Merge:true;Borders:[" + bordoTop + "," + bordoBottom + ", Right:medium]";
                    Style.RangeStyle(rng, style);
                    rng.Value = info["DesInformazione"].ToString();
                }
                else
                {
                    Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaTitoloInfo], _ws.Cells[rigaAttiva, colonnaTitoloInfo + VisParametro - 1]];                    
                    style += "Borders:[insidev:thin,right:medium," + bordoTop + "," + bordoBottom + "]";
                    Style.RangeStyle(rng, style);

                    object[] valori = new object[VisParametro];
                    valori[0] = info["DesInformazione"];
                    valori[1] = info["DesInformazioneBreve"];
                    
//TODO creare struttura per COLONNA PARAMETRO                    
                    if (Struttura.visParametro) 
                        valori[2] = "";
                    
                    string nome = "";

                    if (info["Selezione"].ToString() != "0") 
                    {
                        nome = (info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"]) + Simboli.UNION + "SEL" + info["Selezione"];
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

        private void CreateNomiCelle(DataView informazioni)
        {
            int rigaAttiva = _rigaAttiva;
            foreach (DataRowView info in informazioni)
            {
                int oraAttiva = _colonnaInizio;
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                    {
                        bool isVisibleData0H24 = giorno == _dataInizio && Struttura.visData0H24;

                        for (int i = 0; i < oreGiorno; i++)
                        {
                            if (i == 0 && isVisibleData0H24) 
                                _nomiDefiniti.Add(info["SiglaEntita"] + Simboli.UNION + info["SiglaInformazione"] + Simboli.UNION + "DATA0" + Simboli.UNION + "H24", rigaAttiva, oraAttiva++);

                            _nomiDefiniti.Add(info["SiglaEntita"] + Simboli.UNION + info["SiglaInformazione"] + Simboli.UNION + suffissoData + Simboli.UNION + "H" + (i + 1), rigaAttiva, oraAttiva++);
                        }
                        return true;
                    });
                rigaAttiva++;
            }
        }

        private void InsertValoriCelle(DataView informazioni)
        {
            int rigaAttiva = _rigaAttiva;
            foreach (DataRowView info in informazioni)
            {
                int oraAttiva = _colonnaInizio;
                int x = 0,y = 0;
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    bool isVisibleData0H24 = giorno == _dataInizio && Struttura.visData0H24;
                    object[,] values = new object[informazioni.Count, oreGiorno];

                    for (int i = 0; i < oreGiorno; i++)
                    {
                        if (i == 0 && isVisibleData0H24)
                        {
                            values[x, y++] = "";
                            oraAttiva++;
                        }

                        if (!(info["ValoreDefault"] is DBNull))
                            values[x, y] = double.Parse(info["ValoreDefault"].ToString().Replace('.', ','));
                        else
                            if (info["FormulaInCella"].Equals("0"))//1")) TODO ripristinare a 0
                                values[x, y] = PreparaFormula(info, suffissoData, i + 1);

                        y++;
                        oraAttiva++;
                    }
                    return true;
                });
                x++;
                rigaAttiva++;
            }
        }

        private void InsertNomiValori(DataView informazioni)
        {
            int rigaAttiva = _rigaAttiva;
            int colonnaInizio = _colonnaInizio;

            foreach (DataRowView info in informazioni)
            {
                string siglaEntitaInfo = (info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"]).ToString();
                string bordoTop = "Top:" + (informazioni[0] == info || info["InizioGruppo"].ToString() == "1" ? "medium" : "thin");
                string bordoBottom = (informazioni[informazioni.Count - 1] == info ? "Bottom:medium" : "");
                bool grassetto = info["Grassetto"].ToString() == "1";
                int backColor = (info["BackColor"] is DBNull ? 0 : (int)info["BackColor"]);
                backColor = backColor == 0 || backColor == 2 ? (info["Editabile"].ToString() == "1" ? 15 : 48) : backColor;

                string formula = "";// PreparaFormula(info);

                //proprietà di stile comuni
                string style = "FontSize=" + info["FontSize"] + ";BackColor=" + backColor + ";"
                    + "ForeColor=" + info["ForeColor"] + ";Visible=" + info["Visibile"] + ";";

                //scrivo i nomi per le ore e personalizzo lo stile dove necessario
                int colonnaAttiva = _colonnaInizio - 1;
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
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
//TODO CONTROLLARE              _ws.Cells[rigaAttiva, colonnaAttiva].Formula = "=SUM($E$25:$FP$25)";
                            }
                        }
                        int oreRealiGiorno = oreGiorno - (isVisibleData0H24 ? 1 : 0);
                        for (int i = 1; i <= oreRealiGiorno; i++)
                        {
                            int oraAttiva = colonnaAttiva + i;

                            _ws.Cells[rigaAttiva, oraAttiva].Name = (info["SiglaEntita"] + Simboli.UNION + info["SiglaInformazione"] + Simboli.UNION + suffissoData + Simboli.UNION + "H" + (i));
                            style = "ForeColor:" + info["ForeColor"] + ";BackColor" + backColor + ";Bold:" + grassetto + ";NumberFormat:" + info["Formato"];
                                //+ "Borders:[InsideV:thin," + bordoTop + "," + bordoBottom + (i == oreRealiGiorno ? ",Right:medium" : "") + "]";
                            Style.RangeStyle(_ws.Cells[rigaAttiva, oraAttiva], style);

                            if (!(info["ValoreDefault"] is DBNull))
                            {
                                _ws.Cells[rigaAttiva, oraAttiva].FormulaR1C1 = double.Parse(info["ValoreDefault"].ToString().Replace('.', ','));
                            }
                            else
                            {
                                if (info["FormulaInCella"].Equals("0"))//")) TODO ripristinare a 0!!!!
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

                        return true;
                    });
                rigaAttiva++;
            }
            rigaAttiva++;
        }

        private string PreparaFormula(DataRowView info, string suffissoData, int ora)
        {
            /*TEST CONDITION*/
            //if (info["FormulaInCella"].Equals("1"))
            if(!(info["Formula"] is DBNull && info["Funzione"] is DBNull))
            {
            /*END*/
                string formula = info["Formula"].ToString();
                if (formula == "")
                    formula = info["Funzione"].ToString().Replace("%SHEET%", _nomeFoglio).Replace("%ENTITA%", info["SiglaEntita"].ToString());
                else
                {
                    string[] parametri = info["FormulaParametro"].ToString().Split(',');
                    formula = Regex.Replace(formula, @"%P\d+%", delegate(Match m)
                        {
                            int n = int.Parse(Regex.Match(m.Value, @"\d+").Value);

                            string nome = info["SiglaEntita"] + Simboli.UNION + parametri[n - 1];
                            if(nome.EndsWith("[-1]"))
                                nome += Simboli.UNION + (ora == 1 ? "DATA0" + Simboli.UNION + "H24" : suffissoData + Simboli.UNION + "H" + (ora - 1));
                            else
                                nome += Simboli.UNION + suffissoData + "H" + ora;

                            Tuple<int, int> coordinate = _nomiDefiniti[nome];
                            //string s = info["SiglaEntita"] + Simboli.UNION + parametri[n - 1].Replace("[-1]","");
                            //return s.ToUpperInvariant() + Simboli.UNION + (parametri[n - 1].EndsWith("[-1]") ? "%DATA-1%" : "%DATA%");

                            return "R" + coordinate.Item1 + "C" + coordinate.Item2;
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