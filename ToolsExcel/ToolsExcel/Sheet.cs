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

namespace Iren.FrontOffice.Tools
{
    class Sheet<T> : CommonFunctions
    {
        #region Strutture

        public struct Struttura
        {            
            public const int COL_BLOCK = 5,
                ROW_BLOCK = 6, 
                ROW_GOTO = 3;
        }
        public struct Col
        {
            public struct Width
            {
                public const double EMPTY = 1,
                CATEGORIA_ORA = 8.8,
                CATEGORIA_ENTITA = 3,
                CATEGORIA_INFORMAZIONE = 28,
                RIEPILOGO = 9;
            }
        }
        public struct Row 
        {
            public struct Height
            {
                public const double NORMAL = 15,
                EMPTY = 5;
            }
        }
        public struct Simboli
        {
            public const string UNION = ".";
        }

        #endregion

        #region Variabili

        Worksheet _ws;
        Dictionary<string, object> _config;
        Dictionary<string, NamedRange> _ranges = new Dictionary<string, NamedRange>();
        
        #endregion

        public Sheet(T categoria)
        {            
            Type t = categoria.GetType();
            PropertyInfo p = t.GetProperty("Base");
            _ws = (Worksheet) p.GetValue(categoria, null);

            FieldInfo f = t.GetField("config");
            _config = (Dictionary<string,object>)f.GetValue(categoria);
            
            StdStyles();
        }

        //private U getCategoria<U>(T cat)
        //{
        //    return (U)Convert.ChangeType(cat, typeof(U));
        //}

        public void AddNamedRange(Excel.Range rng, string name, string internalName = "")
        {
            if (internalName == "")
                internalName = name;

            try
            {
                _ranges.Add(internalName, (NamedRange)_ws.Controls[name]);
            }
            catch
            {
                _ranges.Add(internalName, _ws.Controls.AddNamedRange(rng, name));
            }
        }

        private void SetAllBorders(Excel.Style s, int colorIndex, Excel.XlBorderWeight weight)
        {
            s.Borders.ColorIndex = 1;
            s.Borders.Weight = weight;
            s.Borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            s.Borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        }

        public static void RangeStyle(Excel.Range rng, string style)
        {
            MatchCollection paramsList = Regex.Matches(style, @"\w+[=][^;=]+");
            foreach (Match par in paramsList)
            {
                string[] keyVal = Regex.Split(par.Value, "[=]");
                if (keyVal.Length != 2)
                    throw new InvalidExpressionException("The provided list of parameters is invalid.");

                switch (keyVal[0].ToLowerInvariant())
                {
                    case "merge":
                        rng.MergeCells = Regex.IsMatch(keyVal[1], "true|1", RegexOptions.IgnoreCase);
                        break;
                    case "bold":
                        rng.Font.Bold = Regex.IsMatch(keyVal[1], "true|1", RegexOptions.IgnoreCase);
                        break;
                    case "fontsize":
                        double size;
                        if (!double.TryParse(keyVal[1], out size))
                            size = 10.0;
                        rng.Font.Size = size;
                        break;
                    case "align":
                        string align = "xlHAlign" + Regex.Replace(keyVal[1], @"Center|Across|Selection|Distributed|Fill|General|Justify|Left|Right",delegate(Match m)
                            {
                                string v = m.ToString();
	                            return char.ToUpper(v[0]) + v.Substring(1);
                            }, RegexOptions.IgnoreCase);

                        rng.HorizontalAlignment = (Excel.XlHAlign)Enum.Parse(typeof(Excel.XlHAlign), align);
                        rng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        break;
                    case "numberformat":
                        rng.NumberFormat = keyVal[1];
                        break;
                    case "forecolor":
                        rng.Font.ColorIndex = int.Parse(keyVal[1]);
                        break;
                    case "backcolor":
                        rng.Interior.ColorIndex = int.Parse(keyVal[1]);
                        break;
                    case "backpattern":
                        string backPattern = "xlPattern" + Regex.Replace(keyVal[1], "Vertical|Up|None|Horizontal|Gray|Down|Automatic|Solid|Checker|Semi|Light|Grid|Criss|Cross|Linear|Gradient|Rectangular",delegate(Match m)
                            {
                                string v = m.ToString();
	                            return char.ToUpper(v[0]) + v.Substring(1);
                            }, RegexOptions.IgnoreCase);

                        rng.Interior.Pattern = (Excel.XlPattern)Enum.Parse(typeof(Excel.XlPattern), backPattern);
                        break;
                    case "borders":
                        MatchCollection borders = Regex.Matches(keyVal[1], @"(Top|Left|Bottom|Right|InsideH|InsideV)(:\w*)?", RegexOptions.IgnoreCase);
                        rng.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                        foreach (Match border in borders)
                        {
                            string[] b = border.Value.Split(':');

                            Excel.XlBordersIndex index = Excel.XlBordersIndex.xlEdgeTop;
                            Excel.XlBorderWeight weight = Excel.XlBorderWeight.xlThin;
                            switch(b[0].ToLowerInvariant())
                            {
                                case "top":
                                    index = Excel.XlBordersIndex.xlEdgeTop;
                                    break;
                                case "left":
                                    index = Excel.XlBordersIndex.xlEdgeLeft;
                                    break;
                                case "bottom":
                                    index = Excel.XlBordersIndex.xlEdgeBottom;
                                    break;
                                case "right":
                                    index = Excel.XlBordersIndex.xlEdgeRight;
                                    break;
                                case "insideh":
                                    index = Excel.XlBordersIndex.xlInsideHorizontal;
                                    break;
                                case "insidev":
                                    index = Excel.XlBordersIndex.xlInsideVertical;
                                    break;
                            }
                            if (b.Length == 2)
                            {
                                switch (b[1].ToLowerInvariant())
                                {
                                    case "thick":
                                        weight = Excel.XlBorderWeight.xlThick;
                                        break;
                                    case "thin":
                                        weight = Excel.XlBorderWeight.xlThin;
                                        break;
                                    case "medium":
                                        weight = Excel.XlBorderWeight.xlMedium;
                                        break;
                                    case "hairline":
                                        weight = Excel.XlBorderWeight.xlHairline;
                                        break;
                                }
                            }
                            rng.Borders[index].LineStyle = Excel.XlLineStyle.xlContinuous;
                            rng.Borders[index].Weight = weight;
                        }
                        break;
                    case "orientation":
                        string orientation = "xl" + Regex.Replace(keyVal[1], "Downward|Horizontal|Upward|Vertical", delegate(Match m)
                            {
                                string v = m.ToString();
                                return char.ToUpper(v[0]) + v.Substring(1);
                            }, RegexOptions.IgnoreCase);

                        rng.Orientation = (Excel.XlOrientation)Enum.Parse(typeof(Excel.XlOrientation), orientation);
                        break;
                    case "visible":
                        rng.EntireRow.Hidden = Regex.IsMatch(keyVal[1], "false|0", RegexOptions.IgnoreCase);
                        break;
                }
            }
        }

        private void StdStyles()
        {
            Excel.Style gotoBar;
            try 
            {
                gotoBar = Globals.ThisWorkbook.Styles["gotoBarStyle"];
            }
            catch 
            {
                gotoBar = Globals.ThisWorkbook.Styles.Add("gotoBarStyle");
                gotoBar.Font.Bold = false;
                gotoBar.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                gotoBar.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                gotoBar.Interior.ColorIndex = 15;
            }

            Excel.Style navBar;
            try
            {
                navBar = Globals.ThisWorkbook.Styles["navBarStyle"];
            }
            catch
            {
                navBar = Globals.ThisWorkbook.Styles.Add("navBarStyle");
                navBar.Font.Bold = true;
                navBar.Font.Size = 7;
                navBar.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                navBar.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                navBar.Interior.ColorIndex = 2;
                SetAllBorders(navBar, 1, Excel.XlBorderWeight.xlThin);
            }

            Excel.Style titleBar;
            try
            {
                titleBar = Globals.ThisWorkbook.Styles["titleBarStyle"];
            }
            catch
            {
                titleBar = Globals.ThisWorkbook.Styles.Add("titleBarStyle");
                titleBar.Font.Bold = true;
                titleBar.Font.Size = 16;
                titleBar.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                titleBar.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                titleBar.Interior.ColorIndex = 37;                
                SetAllBorders(titleBar, 1, Excel.XlBorderWeight.xlMedium);
            }

            Excel.Style dateBar;
            try
            {
                dateBar = Globals.ThisWorkbook.Styles["dateBarStyle"];
            }
            catch
            {
                dateBar = Globals.ThisWorkbook.Styles.Add("dateBarStyle");
                dateBar.Font.Bold = true;
                dateBar.Font.Size = 10;
                dateBar.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                dateBar.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                dateBar.NumberFormat = "dddd d mmmm yyyy";
                dateBar.Interior.ColorIndex = 15;
                SetAllBorders(dateBar, 1, Excel.XlBorderWeight.xlMedium);
            }

            Excel.Style chartsBar;
            try
            {
                chartsBar = Globals.ThisWorkbook.Styles["chartsBarStyle"];
            }
            catch
            {
                chartsBar = Globals.ThisWorkbook.Styles.Add("chartsBarStyle");
                chartsBar.Font.Size = 10;
                chartsBar.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                chartsBar.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                chartsBar.NumberFormat = "dddd d mmmm yyyy";
                chartsBar.Interior.ColorIndex = 2;
                chartsBar.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            }
        }

        public void Clear()
        {
            int dataOreTot = ThisWorkbook.Parameters.DATA_ORE_TOT;            
            
            _ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            
            Excel.Range workingRange = _ws.Range[_ws.Cells[1, 2], _ws.Cells[1, dataOreTot * 3]];

            workingRange.EntireColumn.Delete();
            _ws.Rows["1:1000"].EntireRow.Hidden = false;
            _ws.Rows["1:1000"].Font.Size = 10;
            _ws.Rows["1:1000"].Font.Name = "Verdana";
            _ws.Rows["1:1000"].RowHeight = Row.Height.NORMAL;

            _ws.Columns["A:A"].ColumnWidth = Col.Width.EMPTY;
            _ws.Range[_ws.Rows[1], _ws.Rows[Struttura.ROW_BLOCK - 1]].RowHeight = Row.Height.EMPTY;
            _ws.Rows[Struttura.ROW_GOTO].RowHeight = Row.Height.NORMAL;

            _ws.Activate();
            _ws.Application.ActiveWindow.FreezePanes = false;
            _ws.Range[_ws.Cells[Struttura.ROW_BLOCK, Struttura.COL_BLOCK], _ws.Cells[Struttura.ROW_BLOCK, Struttura.COL_BLOCK]].Select();
            _ws.Application.ActiveWindow.ScrollColumn = 1;
            _ws.Application.ActiveWindow.ScrollRow = 1;
            _ws.Application.ActiveWindow.FreezePanes = true;

            string gotoBarRangeName = _config["SiglaCategoria"] + Simboli.UNION + "GOTO_BAR";
            Excel.Range rng = _ws.Range[_ws.Cells[2, 2], _ws.Cells[Struttura.ROW_BLOCK - 2, 
                Struttura.COL_BLOCK + dataOreTot - 1]];
            AddNamedRange(rng, gotoBarRangeName, "gotoBarRange");
            _ranges["gotoBarRange"].Style = "gotoBarStyle";
            _ranges["gotoBarRange"].BorderAround2(Weight: Excel.XlBorderWeight.xlMedium, Color: 1);
        }

        public void LoadStructure()
        {
            DataView dvCE = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.CATEGORIAENTITA].DefaultView;
            DataView dvEP = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAPROPRIETA].DefaultView;

            dvCE.RowFilter = "SiglaCategoria = '" + _config["SiglaCategoria"] + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";
            DateTime dataInizio = (DateTime)_config["DataInizio"];
            int intervalloGiorni = (int)_config["IntervalloGiorni"];

            InitBarraNavigazione(dvCE);

            //int intervalloOre;
            int rigaAttiva = Struttura.ROW_BLOCK + 1;

            foreach (DataRowView entita in dvCE)
            {
                string siglaEntita = ""+entita["SiglaEntita"];
                dvEP.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta LIKE '%GIORNI_STRUTTURA'";
                DateTime dataFine;
                if (dvEP.Count > 0)
                    dataFine = dataInizio.AddDays(double.Parse("" + dvEP[0]["Valore"]));
                else
                    dataFine = dataInizio.AddDays(intervalloGiorni);

                InitBloccoEntita(entita, dataInizio, dataFine, ref rigaAttiva);                

            }
        }

        private void InitBloccoEntita(DataRowView entita, DateTime dataInizio, DateTime dataFine, ref int rigaAttiva)
        {
            DataView grafici = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAGRAFICO].DefaultView;
            grafici.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

            DataView informazioni = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAINFORMAZIONE].DefaultView;
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

            int colonnaInizio = Struttura.COL_BLOCK;
            int intervalloOre = CommonFunctions.GetOreIntervallo(dataInizio, dataFine);

            //titolo + data
            InsertTitoloEntita(entita, ref rigaAttiva, colonnaInizio, dataInizio, dataFine);
            //grafici
            InsertGraficiEntita(grafici, ref rigaAttiva, colonnaInizio, intervalloOre);
            //informazioni + ore
            InsertInformazioniEntita(informazioni, ref rigaAttiva, colonnaInizio - 2, intervalloOre);
        }

        private void InsertTitoloEntita(DataRowView entita, ref int rigaAttiva, int colonnaInizio, DateTime dataInizio, DateTime dataFine)
        {
            for (var giorno = dataInizio; giorno <= dataFine; giorno = giorno.AddDays(1))
            {
                int oreGiorno = CommonFunctions.GetOreGiorno(giorno);
                string suffissoData = CommonFunctions.GetSuffissoData(dataInizio, giorno);

                string rangeTitolo = entita["SiglaEntita"] + Simboli.UNION + "T" + Simboli.UNION + suffissoData;
                string rangeData = entita["SiglaEntita"] + Simboli.UNION + suffissoData;

                Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaInizio],
                        _ws.Cells[rigaAttiva, colonnaInizio + oreGiorno - 1]];

                AddNamedRange(rng, rangeTitolo);

                _ranges[rangeTitolo].Merge();
                _ranges[rangeTitolo].Style = "titleBarStyle";
                _ranges[rangeTitolo].Value = entita["DesEntita"].ToString().ToUpperInvariant();
                _ranges[rangeTitolo].RowHeight = 25;

                rng = _ws.Range[_ws.Cells[rigaAttiva + 1, colonnaInizio],
                        _ws.Cells[rigaAttiva + 1, colonnaInizio + oreGiorno - 1]];

                AddNamedRange(rng, rangeData);

                _ranges[rangeData].Merge();
                _ranges[rangeData].Style = "dateBarStyle";
                _ranges[rangeData].Value = giorno.ToString("MM/dd/yyyy");
                _ranges[rangeData].RowHeight = 20;       

                colonnaInizio += oreGiorno;
            }
            rigaAttiva++;
        }
        private void InsertGraficiEntita(DataView grafici, ref int rigaAttiva, int colonnaInizio, int intervalloOre)
        {            
            int i = 1;
            foreach (DataRowView grafico in grafici)
            {
                string graficoRange = grafico["SiglaEntita"] + Simboli.UNION + "GRAFICO" + (grafici.Count > 1 ? ""+i++ : "");

                Excel.Range rng = _ws.Range[_ws.Cells[++rigaAttiva, colonnaInizio],
                    _ws.Cells[rigaAttiva, colonnaInizio + intervalloOre - 1]];

                AddNamedRange(rng, graficoRange);

                _ranges[graficoRange].Merge();
                _ranges[graficoRange].Style = "chartsBarStyle";
                _ranges[graficoRange].RowHeight = 200;
            }
            rigaAttiva++;
        }
        private void InsertInformazioniEntita(DataView informazioni, ref int rigaAttiva, int colonnaInizio, int intervalloOre)
        {
            int count = 0;
            foreach (DataRowView info in informazioni)
            {
                string siglaEntitaInfo = (info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"]).ToString();
                string siglaInfo = info["SiglaInformazione"].ToString();
                string bordoTop = "Top:" + (informazioni[0] == info ||  info["InizioGruppo"].ToString() == "1" ? "medium" : "thin");
                string bordoBottom = "Bottom:" + (informazioni[informazioni.Count - 1] == info ? "medium" : "thin");
                bool grassetto = info["Grassetto"].ToString() == "1";
                int backColor = (info["Editabile"].ToString() == "1" ? 15 : 48);
                string style = "";

                if (info["SiglaTipologiaInformazione"].ToString() == "TITOLO2")
                {
                    Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaInizio], _ws.Cells[rigaAttiva, colonnaInizio + intervalloOre - 1]];
                    style = "Merge=true;Bold=" + (grassetto) + ";FontSize=" + info["FontSize"] + ";BackColor=" + info["BackColor"] + ";"
                        + "ForeColor=" + info["ForeColor"] + ";Borders=[" + bordoTop + "," + bordoBottom + "];Visible=" + info["Visibile"];
                    
                    RangeStyle(rng, style);
                    rng.Value = info["DesInformazione"].ToString();
                }
                else
                {
                    Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaInizio], _ws.Cells[rigaAttiva, colonnaInizio + 1]];
                    style += "Bold=false;FontSize=9;BackColor=" + backColor + ";"
                        + "Borders=[insidev:thin,right:medium," + bordoTop + "," + bordoBottom + "];Visible=" + info["Visibile"];
                        
                    RangeStyle(rng, style);

                    object[] valori = new object[2];
                    valori[0] = info["DesInformazione"];
                    valori[1] = info["DesInformazioneBreve"];
                    string nome = "";

                    if (siglaEntitaInfo == "UP_BUS")
                    {
                        nome = siglaEntitaInfo + Simboli.UNION + info["SiglaInformazione"] + Simboli.UNION + "DATA0.H24";
                        _ws.Cells[rigaAttiva, colonnaInizio + 1].Interior.ColorIndex = 2;

                        if (info["SiglaInformazione"].ToString() == "VOL_INVASO")
                        {
                            valori[0] = info["DesInformazione"] + " " + info["DesInformazioneBreve"];
                            valori[1] = 160;
                        }
                        else if (info["SiglaInformazione"].ToString() == "TEMP_PROG5")
                        {
                            valori[1] = "=SUM($E$25:$FP$25)";
                        }
                        else
                        {
                            valori[1] = info["DesInformazioneBreve"];
                        }
                    } else if (info["Selezione"].ToString() != "0") 
                    {
                        nome = siglaEntitaInfo + Simboli.UNION + "SEL" + info["Selezione"];
                    }


                    //scrivo i valori sulle celle
                    rng.Value = valori;
                    //scrivo i nomi dove necessario
                    _ws.Cells[rigaAttiva, colonnaInizio].Name = siglaEntitaInfo + Simboli.UNION + info["SiglaInformazione"];
                    if(nome != "")
                        _ws.Cells[rigaAttiva, colonnaInizio + 1].Name = nome;
                    //cambio impostazioni per la seconda cella
                    _ws.Cells[rigaAttiva, colonnaInizio + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                rigaAttiva++;
            }
        }


        private void InitBarraNavigazione(DataView entita)
        {
            object[] descrizioni = new object[entita.Count];
            int i = -1;
            foreach (DataRowView e in entita)
            {
                descrizioni[++i] = e["DesEntitaBreve"];
                _ws.Cells[Struttura.ROW_GOTO, Struttura.COL_BLOCK + i].Name = e["siglaEntita"] + Simboli.UNION + "GOTO";
            }

            Excel.Range rng = _ws.Range[_ws.Cells[Struttura.ROW_GOTO, Struttura.COL_BLOCK],
                _ws.Cells[Struttura.ROW_GOTO, Struttura.COL_BLOCK + i]];
            string gotoMenuRangeName = _config["SiglaCategoria"] + Simboli.UNION + "GOTO_MENU";

            AddNamedRange(rng, gotoMenuRangeName, "gotoMenuRange");
            _ranges["gotoMenuRange"].Value = descrizioni;
            _ranges["gotoMenuRange"].Style = "navBarStyle";            
        }
    }
}