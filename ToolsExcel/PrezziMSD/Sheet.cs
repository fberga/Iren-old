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
            int rigaAttiva = _rigaAttiva;
            //sono alla prima riga vuota dopo le informazioni
            DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPROPRIETA].DefaultView;
            entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'MSD_ACCENSIONE'";

            int colonnaTitoloInfo = _colonnaInizio - _visParametro;

            if (entitaProprieta.Count > 0)
            {
                Excel.Range rng = _ws.Range[new Range(_rigaAttiva, colonnaTitoloInfo,1, 2).ToString()];
                rng.Value = new object[] {"Accensione",""};
                Style.RangeStyle(rng, backColor: 36, align: Excel.XlHAlign.xlHAlignCenter, borders: "top:medium,right:medium,bottom:medium,left:medium,insidev:thin", numberFormat: "#,##0;-#,##0;-");
                Style.RangeStyle(rng.Cells[1], align: Excel.XlHAlign.xlHAlignLeft);

                _newNomiDefiniti.AddName(_rigaAttiva++, siglaEntita, "ACCENSIONE", Date.GetSuffissoData(_dataInizio));

            }

            entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'MSD_CAMBIO_ASSETTO'";

            if (entitaProprieta.Count > 0)
            {
                Excel.Range rng = _ws.Range[new Range(_rigaAttiva, colonnaTitoloInfo, 1, 2).ToString()];
                rng.Value = new object[] { "Cambio Assetto", "" };
                Style.RangeStyle(rng, backColor: 35, align: Excel.XlHAlign.xlHAlignCenter, borders: "top:medium,right:medium,bottom:medium,left:medium,insidev:thin", numberFormat: "#,##0;-#,##0;-");
                Style.RangeStyle(rng.Cells[1], align: Excel.XlHAlign.xlHAlignLeft);
                _newNomiDefiniti.AddName(_rigaAttiva++, siglaEntita, "CAMBIO_ASSETTO", Date.GetSuffissoData(_dataInizio));
            }

            if (rigaAttiva != _rigaAttiva)
            {
                //estendo titolo verticale
                Range rng = new Range(rigaAttiva - 1, colonnaTitoloInfo - 1, _rigaAttiva - rigaAttiva + 1);
                _ws.Range[rng.Rows[1, rng.Rows.Count].ToString()].Style = "titoloVertStyle";
                _ws.Range[rng.ToString()].Merge();
            }

        }

        public override void CaricaInformazioni(bool all)
        {
            base.CaricaInformazioni(all);

            //carico le informazioni giornaliere
            DataView datiApplicazioneD = DataBase.DB.Select(DataBase.SP.APPLICAZIONE_INFORMAZIONE_D, "@SiglaCategoria=" + _siglaCategoria + ";@SiglaEntita=ALL;@DateFrom=" + DataBase.DataAttiva.ToString("yyyyMMdd") + ";@DateTo=" + DataBase.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("yyyyMMdd") + ";@Tipo=1;@All=" + (all ? "1" : "0")).DefaultView;

            foreach (DataRowView dato in datiApplicazioneD)
            {
                //DateTime giorno = DateTime.ParseExact(dato["Data"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                //sono nel caso DATA0H24
                //if (giorno < DataBase.DataAttiva)
                //{
                //    Range rng = _newNomiDefiniti.Get(dato["SiglaEntita"], dato["SiglaInformazione"], Date.GetSuffissoData(DataBase.DataAttiva.AddDays(-1)), Date.GetSuffissoOra(24));
                //    _ws.Range[rng.ToString()].Value = dato["H24"];
                //}
                //else
                //{
                    //int col = Struct.tipoVisualizzazione == "O" ? _newNomiDefiniti.GetColFromDate(giorno) : _newNomiDefiniti.GetFirstCol();
                    //int dayOffset = Struct.tipoVisualizzazione == "O" ? _newNomiDefiniti.GetDayOffset(giorno) : _newNomiDefiniti.GetColOffset();
                    //int row = _newNomiDefiniti.GetRowByName(dato["SiglaEntita"], dato["SiglaInformazione"], Struct.tipoVisualizzazione == "O" ? "" : Date.GetSuffissoData(giorno));

                Range rng = new Range(_newNomiDefiniti.GetRowByName(dato["SiglaEntita"], dato["SiglaInformazione"], Date.GetSuffissoData(dato["Data"].ToString())), _newNomiDefiniti.GetFirstCol() - 1);

                    _ws.Range[rng.ToString()].Value = dato["Valore"];
                //}
            }

        }
    }
}
