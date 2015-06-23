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

        DefinedNames _definedNamesMercatoAttivo = new DefinedNames(Simboli.Mercato);

        #endregion

        #region Costruttori

        public Sheet(Excel.Worksheet ws) 
            : base(ws) {}
        
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

        protected override void CaricaInformazioniEntita(DataView datiApplicazione)
        {
            foreach (DataRowView dato in datiApplicazione)
            {
                DateTime giorno = DateTime.ParseExact(dato["Data"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                //sono nel caso DATA0H24
                if (giorno < DataBase.DataAttiva)
                {
                    Range rng = _definedNames.Get(dato["SiglaEntita"], dato["SiglaInformazione"], Date.GetSuffissoData(DataBase.DataAttiva.AddDays(-1)), Date.GetSuffissoOra(24));
                    _ws.Range[rng.ToString()].Value = dato["H24"];
                }
                else
                {
                    Range rng = _definedNames.Get(dato["SiglaEntita"], dato["SiglaInformazione"], Date.GetSuffissoData(giorno)).Extend(colOffset: Date.GetOreGiorno(giorno));
                    List<object> o = new List<object>(dato.Row.ItemArray);
                    o.RemoveRange(o.Count - 3, 3);
                    _ws.Range[rng.ToString()].Value = o.ToArray();

                    if (giorno == DataBase.DataAttiva && Regex.IsMatch(dato["SiglaInformazione"].ToString(), @"RIF\d+"))
                    {
                        Selection s = _definedNames.GetSelectionByRif(rng);
                        s.ClearSelections(_ws);
                        s.Select(_ws, int.Parse(o[0].ToString().Split('.')[0]));
                    }
                }
            }
        }
        
        #endregion
    }
}
