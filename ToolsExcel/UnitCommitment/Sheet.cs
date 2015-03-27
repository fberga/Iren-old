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

        protected override void InsertPersonalizzazioni(object siglaEntita)
        {
            //DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            //informazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
            //informazioni.Sort = "Ordine";

            //_ws.Columns[3].Font.Size = 9;

            ////da classe base _dataInizio e _dataFine sono corretti
            //CicloGiorni((oreGiorno, suffissoData, giorno) => 
            //{   
            //    object siglaInfo = DefinedNames.GetName(informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"]);
            //    Tuple<int, int> primaRiga = _nomiDefiniti[DefinedNames.GetName(siglaInfo, informazioni[0]["SiglaInformazione"], suffissoData)].Last();
            //    Excel.Range rng = _ws.Range[_ws.Cells[primaRiga.Item1, primaRiga.Item2 + 1], _ws.Cells[primaRiga.Item1 + informazioni.Count - 1, primaRiga.Item2 + 1]];
            //    rng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
            //    rng.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            //    rng.Columns[1].ColumnWidth = _cell.Width.jolly1;
            //    int r = 0;
            //    foreach (DataRowView info in informazioni)
            //    {
            //        siglaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
            //        _nomiDefiniti.Add(DefinedNames.GetName(siglaInfo, "NOTE", suffissoData), primaRiga.Item1 + r, primaRiga.Item2 + 1, true, true, false);
            //        r++;
            //    }

            //});
        }
        public override void CaricaInformazioni(bool all)
        {
            base.CaricaInformazioni(all);

            try
            {
                if (DataBase.OpenConnection())
                {
                    string start = DataBase.DataAttiva.ToString("yyyyMMdd");
                    string end = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("yyyyMMdd");

                    DataView note = DataBase.DB.Select(DataBase.SP.APPLICAZIONE_NOTE, "@SiglaEntita=ALL;@DateFrom="+start+";@DateTo="+end).DefaultView;

                    foreach (DataRowView nota in note)
                    {
                        Tuple<int, int> cella = _nomiDefiniti[DefinedNames.GetName(nota["SiglaEntita"], "NOTE", Date.GetSuffissoData(nota["Data"].ToString()))][0];
                        Excel.Range rng = _ws.Cells[cella.Item1, cella.Item2];

                        rng.Value = nota["Note"];
                    }

                    DataBase.CloseConnection();
                } 
            }
            catch (Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "CaricaInformazioni Custom UnitComm [all = " + all + "]: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

    }
}
