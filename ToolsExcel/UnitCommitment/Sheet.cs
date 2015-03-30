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

        protected override void InsertPersonalizzazioni()
        {
            //da classe base il filtro è corretto
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;

            _ws.Columns[3].Font.Size = 9;

            int col = _newNomiDefiniti.GetFirstCol();
            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            int row = _newNomiDefiniti.GetRowByName(siglaEntita, informazioni[0]["SiglaInformazione"], Struct.tipoVisualizzazione == "O" ? "" : Date.GetSuffissoData(_dataInizio));

            Excel.Range rngPersonalizzazioni = _ws.Range[GetRange(row, col + 25, informazioni.Count - 1, 0)];

            rngPersonalizzazioni.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
            rngPersonalizzazioni.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            rngPersonalizzazioni.Columns[1].ColumnWidth = _cell.Width.jolly1;

            //da classe base _dataInizio e _dataFine sono corretti
            int i = 0;
            foreach (DataRowView info in informazioni)
            {
                siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                _newNomiDefiniti.AddName(row + i++, siglaEntita, "NOTE", Date.GetSuffissoData(_dataInizio));
            }
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
                        int row = _newNomiDefiniti.GetRowByName(nota["SiglaEntita"], "NOTE", Date.GetSuffissoData(nota["Data"].ToString()));
                        int col = _newNomiDefiniti.GetFirstCol();
                        _ws.Range[GetRange(row, col + 25)].Value = nota["Note"];
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
