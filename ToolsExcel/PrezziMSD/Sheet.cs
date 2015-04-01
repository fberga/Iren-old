using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    class Sheet : Base.Sheet
    {
        public Sheet(Excel.Worksheet ws)
            : base(ws) { }

        protected override void InsertTitoloEntita(object siglaEntita, object desEntita)
        {
            //DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPROPRIETA].DefaultView;
            //entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'MSD_ACCENSIONE'";

            //int colonnaTitoloInfo = _colonnaInizio - _visParametro;

            //if (entitaProprieta.Count > 0)
            //{
            //    //dalla classe base mi trovo in _rigaAttiva = riga del titolo
            //    Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, colonnaTitoloInfo], _ws.Cells[_rigaAttiva, colonnaTitoloInfo + 1]];
            //    rng.Value = new object[] {"Accensione",""};
            //    Style.RangeStyle(rng, "BackColor:36;Align:Center;Borders:[top:medium,right:medium,bottom:medium,left:medium,insidev:thin];NumberFormat:[#,##0;-#,##0;-]");
            //    CicloGiorni((oreGiorno, suffissoData, giorno) =>
            //    {
            //        _nomiDefiniti.Add(DefinedNames.GetName(siglaEntita, "ACCENSIONE", suffissoData), _rigaAttiva, colonnaTitoloInfo + 1, true, true, true);
            //    });
            //}

            //entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'MSD_CAMBIO_ASSETTO'";

            //if (entitaProprieta.Count > 0)
            //{
            //    //dalla classe base mi trovo in _rigaAttiva = riga del titolo
            //    Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva + 1, colonnaTitoloInfo], _ws.Cells[_rigaAttiva + 1, colonnaTitoloInfo + 1]];
            //    rng.Value = new object[] { "Cambio Assetto", "" };
            //    Style.RangeStyle(rng, "BackColor:35;Align:Center;Borders:[top:medium,right:medium,bottom:medium,left:medium,insidev:thin];NumberFormat:[#,##0;-#,##0;-]");
            //    CicloGiorni((oreGiorno, suffissoData, giorno) =>
            //    {
            //        _nomiDefiniti.Add(DefinedNames.GetName(siglaEntita, "CAMBIO_ASSETTO", suffissoData), _rigaAttiva + 1, colonnaTitoloInfo + 1, true, true, true);
            //    });
            //}


            base.InsertTitoloEntita(siglaEntita, desEntita);
        }
    }
}
