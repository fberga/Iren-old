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
        DefinedNames _nomiDefiniti;
        Cell _cell;
        Struttura _struttura;
        
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

            string gotoBarRangeName = _config["SiglaCategoria"] + Simboli.UNION + "GOTO_BAR";
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
            DataView dvEP = LocalDB.Tables[Tab.ENTITAPROPRIETA].DefaultView;
            DataView dvC = LocalDB.Tables[Tab.CATEGORIA].DefaultView;
            DataView dvCE = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;

            dvC.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "'";
            _nomeFoglio = dvC[0]["DesCategoria"].ToString();
            dvCE.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";
            _dataInizio = (DateTime)_config["DataInizio"];

            Clear();
            InitBarraNavigazione(dvCE);

            _rigaAttiva = _struttura.rigaBlock;

            foreach (DataRowView entita in dvCE)
            {
                string siglaEntita = ""+entita["SiglaEntita"];
                dvEP.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                if (dvEP.Count > 0)
                    _dataFine = _dataInizio.AddDays(double.Parse("" + dvEP[0]["Valore"]));
                else
                    _dataFine = _dataInizio.AddDays(_struttura.intervalloGiorni);

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
                //_ws.Cells[_struttura.rigaGoto, _struttura.colBlock + i].Name = e["siglaEntita"] + Simboli.UNION + "GOTO";
                _nomiDefiniti.Add(e["siglaEntita"] + Simboli.UNION + "GOTO", Tuple.Create(_struttura.rigaGoto, _struttura.colBlock + i));
            }

            Excel.Range rng = _ws.Range[_ws.Cells[_struttura.rigaGoto, _struttura.colBlock],
                _ws.Cells[_struttura.rigaGoto, _struttura.colBlock + i]];
            //string gotoMenuRangeName = _config["SiglaCategoria"] + Simboli.UNION + "GOTO_MENU";

            rng.Value = descrizioni;
            rng.Style = "navBarStyle";
        }
        
        private void InitBloccoEntita(DataRowView entita)
        {
            _rigaAttiva++;
            DataView grafici = LocalDB.Tables[Tab.ENTITAGRAFICO].DefaultView;
            grafici.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

            DataView informazioni = LocalDB.Tables[Tab.ENTITAINFORMAZIONE].DefaultView;
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

            DataView formattazione = LocalDB.Tables[Tab.ENTITAINFORMAZIONEFORMATTAZIONE].DefaultView;

            _colonnaInizio = _struttura.colBlock;
            _intervalloOre = GetOreIntervallo(_dataInizio, _dataFine) + (_struttura.visData0H24 ? 1 : 0) + (_struttura.visParametro ? 1 : 0);

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
            //nomi
            CreaNomiCelle(informazioni);
            //valori e formule
            InsertValoriCelle(informazioni);
            //formattazione condizionale
            CreaFormattazioneCondizionale(informazioni, formattazione);
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
                    if (giorno == _dataInizio && _struttura.visData0H24)
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

            object grassetto = "Bold:" + informazioni[0]["Grassetto"];
            string formato = "NumberFormat:[" + informazioni[0]["Formato"] + "]";

            bool primaRigaTitolo2 = informazioni[0]["SiglaTipologiaInformazione"].ToString() == "TITOLO2";
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
                        Style.RangeStyle(allDati, "Style:allDatiStyle;" + grassetto + ";" + formato);
                        allDati.Name = info["SiglaEntita"] + Simboli.UNION + suffissoData + Simboli.UNION + "ALLDATI" + allDatiIndice;
                        allDati.EntireColumn.ColumnWidth = _cell.Width.dato;
                        allDati.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
                        colonnaInizioAllDati += oreGiorno;
                        grassetto = "Bold:" + info["Grassetto"];
                        formato = "NumberFormat:[" + info["Formato"] + "]";
                                
                        return true;
                    });
                    allDatiIndice++;
                    rigaInizioGruppo = rigaAttiva + (info["SiglaTipologiaInformazione"].ToString() == "TITOLO2" ? 1 : 0);
                }
                if (primaRiga && primaRigaTitolo2)
                    rigaInizioGruppo++;

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
                string style = "FontSize:" + info["FontSize"] + ";BackColor:" + backColor + ";"
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
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    bool isVisibleData0H24 = giorno == _dataInizio && _struttura.visData0H24;

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
            int colonnaAttiva = _colonnaInizio;
            string suffissoDataPrec = "DATA0";
            //carico tutti i dati reperibili durante la creazione del foglio
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                int rigaAttiva = _rigaAttiva;
                object[,] values = new object[informazioni.Count, oreGiorno];
                int x = 0;

                foreach (DataRowView info in informazioni)
                {
                    if (!info["SiglaTipologiaInformazione"].Equals("TITOLO2"))
                    {
                        bool isVisibleData0H24 = giorno == _dataInizio && _struttura.visData0H24;

                        int y = 0;
                        for (int i = 0; i < oreGiorno; i++)
                        {
                            if (i == 0 && isVisibleData0H24)
                            {
                                values[x, y] = "";                            
                            }
                            else if (!(info["ValoreDefault"] is DBNull))
                                values[x, y] = double.Parse(info["ValoreDefault"].ToString().Replace('.', ','));
                            else if (info["FormulaInCella"].Equals("1"))
                                values[x, y] = PreparaFormula(info, suffissoDataPrec, suffissoData, i + 1);
                            //else
                            //    values[x, y] = "";
                        
                            y++;
                        }
                    }
                    x++;
                }
                suffissoDataPrec = suffissoData;

                //setto il range
                _ws.Range[_ws.Cells[_rigaAttiva, colonnaAttiva], _ws.Cells[_rigaAttiva + informazioni.Count - 1, colonnaAttiva + oreGiorno - 1]].FormulaR1C1 = values;
                colonnaAttiva += oreGiorno;
                return true;
            });
        }
        private string PreparaFormula(DataRowView info, string suffissoDataPrec, string suffissoData, int ora)
        {
            if(!(info["Formula"] is DBNull && info["Funzione"] is DBNull))
            {
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
                                nome += Simboli.UNION + (ora == 1 ? suffissoDataPrec + Simboli.UNION + "H24" : suffissoData + Simboli.UNION + "H" + (ora - 1));
                            else
                                nome += Simboli.UNION + suffissoData + Simboli.UNION + "H" + ora;
                            nome = nome.Replace("[-1]", "");

                            Tuple<int, int> coordinate = _nomiDefiniti[nome];

                            return "R" + coordinate.Item1 + "C" + coordinate.Item2;
                        }, RegexOptions.IgnoreCase);
                }
                return "=" + formula;
            }
            return "";
        }

        private void CreaFormattazioneCondizionale(DataView informazioni, DataView formattazione)
        {
            foreach (DataRowView info in informazioni)
            {
                string siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"].ToString() : info["SiglaEntitaRif"].ToString();
                formattazione.RowFilter = (info["SiglaEntitaRif"] is DBNull ? "SiglaEntita" : "SiglaEntitaRif") + " = '"+siglaEntita+"' AND SiglaInformazione = '" + info["SiglaInformazione"] + "'";

                foreach (DataRowView format in formattazione)
                {
                    int colonnaInizio = _colonnaInizio;
                    CicloGiorni((oreGiorno,suffissoData, giorno) => 
                    {
                        Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, colonnaInizio], _ws.Cells[_rigaAttiva, colonnaInizio + oreGiorno - 1]];
                        string[] valore = format["Valore"].ToString().Replace("\"","").Split('|');
                        if (!(format["NomeCella"] is DBNull)) 
                        {
                            Tuple<int, int> coordinate = _nomiDefiniti[siglaEntita + Simboli.UNION + format["NomeCella"] + Simboli.UNION + suffissoData + Simboli.UNION + "H1"];
                            string address = _ws.Application.ConvertFormula("R" + coordinate.Item1 + "C" + coordinate.Item2, Excel.XlReferenceStyle.xlR1C1, Excel.XlReferenceStyle.xlA1).Replace("$","");

                            string formula = "";
                            switch((int)format["Operatore"]) {
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
                            if((int)format["BackColor"] != 0)
                                cond.Interior.Color =  format["BackColor"];
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
                        colonnaInizio += oreGiorno;

                        return true;
                    });
                }
                _rigaAttiva++;
            }
            _rigaAttiva++;
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