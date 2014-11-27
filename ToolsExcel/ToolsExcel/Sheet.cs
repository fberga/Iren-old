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

        #endregion

        #region Variabili

        string _siglaCategoria;
        Worksheet _ws;
        
        #endregion

        public Sheet(T categoria)
        {            
            Type t = categoria.GetType();
            PropertyInfo p = t.GetProperty("Base");
            _ws = (Worksheet) p.GetValue(categoria, null);
            FieldInfo f = t.GetField("CATEGORIA");
            _siglaCategoria = f.GetValue(categoria).ToString();
            StdStyles();
        }

        private U getCategoria<U>(T cat)
        {
            return (U)Convert.ChangeType(cat, typeof(U));
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

            NamedRange gotoBarRange;
            try
            {
                gotoBarRange = (NamedRange)_ws.Controls["gotoBarRange"];
            }
            catch
            {
                Excel.Range rng = _ws.Range[_ws.Cells[2, 2], _ws.Cells[Struttura.ROW_BLOCK - 2, Struttura.COL_BLOCK + dataOreTot - 1]];
                gotoBarRange = _ws.Controls.AddNamedRange(rng, "gotoBarRange");
                gotoBarRange.Style = "gotoBarStyle";
                gotoBarRange.BorderAround2(Weight: Excel.XlBorderWeight.xlMedium, Color: 1);
            }
        }

        public void LoadStructure()
        {
            DataView dvCE = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.CATEGORIAENTITA].DefaultView;
            dvCE.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";

            foreach (DataRowView rCE in dvCE)
            {


            }
        }
    }
}