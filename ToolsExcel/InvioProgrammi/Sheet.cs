using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    public class Sheet : Base.Sheet
    {
        #region Variabili

        DefinedNames _definedNamesSheetMercato = new DefinedNames("MSD1");  //non mi interessa sapere il mercato... sono tutti uguali
        string _mercatoPrec = Simboli.GetMercatoPrec();
        Excel.Worksheet _wsMercatoPrec;

        #endregion

        #region Costruttori

        public Sheet(Excel.Worksheet ws) 
            : base(ws) 
        {
            _wsMercatoPrec = Workbook.Sheets[this._mercatoPrec ?? "MSD1"];
        
        }
        
        #endregion

        #region Metodi

        protected override void InsertTitoloVerticale(object desEntita)
        {
            base.InsertTitoloVerticale(desEntita);

            //rimuovo la scritta
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;

            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            Range rngTitolo = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), _struttura.colBlock - _visParametro - 1, informazioni.Count);

            Excel.Range titoloVert = _ws.Range[rngTitolo.ToString()];
            titoloVert.Value = null;
        }

        protected override void InsertInformazioniEntita()
        {
            base.InsertInformazioniEntita();
        }

        protected override void CaricaInformazioniEntita(DataView datiApplicazione)
        {
            base.CaricaInformazioniEntita(datiApplicazione);
            if (_mercatoPrec != null)
            {
                SplashScreen.UpdateStatus("Aggiorno colori");
                DataTable entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA];

                foreach (DataRowView dato in datiApplicazione)
                {
                    DateTime giorno = DateTime.ParseExact(dato["Data"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                    Range rngMercato = _definedNames.Get(dato["SiglaEntita"], dato["SiglaInformazione"], Date.GetSuffissoData(giorno)).Extend(colOffset: Date.GetOreGiorno(giorno));

                    string quarter = Regex.Match(dato["SiglaInformazione"].ToString(), @"Q\d").Value;
                    quarter = quarter == "" ? "Q1" : quarter;

                    var rif =
                        (from r in entita.AsEnumerable()
                         where r["SiglaEntita"].Equals(dato["SiglaEntita"])
                         select new {SiglaEntita = r["Gerarchia"] is DBNull ? r["SiglaEntita"] : r["Gerarchia"], Riferimento = r["Riferimento"]}).First();


                    Range rngMercatoPrec = new Range(_definedNamesSheetMercato.GetRowByName(rif.SiglaEntita, "UM", "T") + 2, _definedNamesSheetMercato.GetColFromName("RIF" + rif.Riferimento, "PROGRAMMA" + quarter)).Extend(rowOffset: Date.GetOreGiorno(giorno));

                    for (int j = 0; j < rngMercatoPrec.Rows.Count; j++)
                        _ws.Range[rngMercato.Columns[j].ToString()].Interior.ColorIndex = _wsMercatoPrec.Range[rngMercatoPrec.Rows[j].ToString()].DisplayFormat.Interior.ColorIndex;
                }
            }
        }
        
        #endregion
    }
}
