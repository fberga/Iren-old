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

        public void AddNamedRange(Excel.Range rng, string name, string internalName)
        {
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

            int intervalloOre;
            int rigaAttiva = Struttura.ROW_BLOCK;
            int colonnaAttiva = 0;
            
            foreach (DataRowView rCE in dvCE)
            {
                string siglaEntita = ""+rCE["SiglaEntita"];
                dvEP.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta LIKE '%GIORNI_STRUTTURA'";
                DateTime dataFine;
                if (dvEP.Count > 0)
                    dataFine = dataInizio.AddDays(double.Parse("" + dvEP[0]["Valore"]));
                else
                    dataFine = dataInizio.AddDays(intervalloGiorni);
                intervalloOre = CommonFunctions.GetOreIntervallo(dataInizio, dataFine);

                rigaAttiva++;
                InitBloccoEntita(rCE, dataInizio, dataFine, rigaAttiva);
                rigaAttiva++;

            }
        }

        private void InitBloccoEntita(DataRowView entita, DateTime inizio, DateTime fine, int rigaAttiva)
        {
            DataView dvEG = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAGRAFICO].DefaultView;
            dvEG.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

            DataView dvEI = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAINFORMAZIONE].DefaultView;
            dvEI.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

            int colonnaInizio = Struttura.COL_BLOCK;

            for (var giorno = inizio; giorno <= fine; giorno = giorno.AddDays(1))
            {
                int oreGiorno = CommonFunctions.GetOreGiorno(giorno);
                string suffissoData = CommonFunctions.GetSuffissoData(inizio, giorno);

                InitTitoloEntita(entita, giorno, ref rigaAttiva, colonnaInizio, oreGiorno, suffissoData);

            }

            

            //titolo
            //data
            
            
            //grafico
            
            
            //informazioni//ore

        }

        private void InitTitoloEntita(DataRowView entita, DateTime giorno, ref int rigaAttiva, int colonnaInizio, int oreGiorno, string suffissoData)
        {
            Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaInizio],
                    _ws.Cells[rigaAttiva, colonnaInizio + oreGiorno - 1]];

            rng.Merge();
            rng.Style = "titleBarStyle";
            rng.Name = entita["SiglaEntita"] + Simboli.UNION + "T" + Simboli.UNION + suffissoData;
            rng.Value = entita["DesEntita"].ToString().ToUpperInvariant();
            rng.RowHeight = 25;

            rng = _ws.Range[_ws.Cells[++rigaAttiva, colonnaInizio],
                    _ws.Cells[rigaAttiva, colonnaInizio + oreGiorno - 1]];

            rng.Merge();
            rng.Style = "dateBarStyle";
            rng.Name = entita["SiglaEntita"] + Simboli.UNION + suffissoData;
            rng.Value = giorno.ToString("MM/dd/yyyy");
            rng.RowHeight = 20;
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
            //_ranges["gotoMenuRange"].BorderAround2(Weight: Excel.XlBorderWeight.xlThin);
        }
    }
}