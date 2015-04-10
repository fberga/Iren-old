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
            //da classe base il filtro è corretto
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;

            _ws.Columns[3].Font.Size = 9;

            int col = _newNomiDefiniti.GetFirstCol();
            siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            int row = _newNomiDefiniti.GetRowByName(siglaEntita, informazioni[0]["SiglaInformazione"], Struct.tipoVisualizzazione == "O" ? "" : Date.GetSuffissoData(_dataInizio));

            Excel.Range rngPersonalizzazioni = _ws.Range[Range.GetRange(row, col + 25, informazioni.Count)];

            rngPersonalizzazioni.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
            rngPersonalizzazioni.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            rngPersonalizzazioni.Columns[1].ColumnWidth = Struct.cell.width.jolly1;

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
                        _ws.Range[Range.GetRange(row, col + 25)].Value = nota["Note"];
                    }
                } 
            }
            catch (Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "CaricaInformazioni Custom UnitComm [all = " + all + "]: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        public override void UpdateData(bool all = true)
        {
            //cancello tutte le NOTE
            if (all)
            {
                DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";

                DateTime dataInizio = DataBase.DataAttiva;
                DateTime dataFine = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni);

                int col = _newNomiDefiniti.GetFirstCol() + 25;

                foreach (DataRowView entita in categoriaEntita)
                {
                    DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
                    informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
                    object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];

                    CicloGiorni(dataInizio, dataFine, (oreGiorno, suffData, g) =>
                    {
                        int row = _newNomiDefiniti.GetRowByName(siglaEntita, informazioni[0]["SiglaInformazione"], suffData);
                        _ws.Range[Range.GetRange(row, col, informazioni.Count)].Value = "";
                    });
                }
            }
            base.UpdateData(all);
        }
    }
}
