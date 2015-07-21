using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    public class Carica : Base.Carica
    {
        DefinedNames _definedNamesSheetMercato = new DefinedNames(Simboli.Mercato);
        Excel.Worksheet _wsMercato;

        public Carica() 
            : base() 
        {
            _wsMercato = Workbook.Sheets[Simboli.Mercato];
        }

        public override bool AzioneInformazione(object siglaEntita, object siglaAzione, object azionePadre, DateTime giorno, object parametro = null)
        {
            bool o = base.AzioneInformazione(siglaEntita, siglaAzione, azionePadre, giorno, parametro);

            string name = DefinedNames.GetSheetName(siglaEntita);
            Sheet s = new Sheet(Workbook.Sheets[name]);
            s.AggiornaColori();

            return o;
        }

        protected override void ScriviCella(Excel.Worksheet ws, DefinedNames definedNames, object siglaEntita, DataRowView info, string suffissoData, string suffissoOra, object risultato, bool saveToDB)
        {
            base.ScriviCella(ws, definedNames, siglaEntita, info, suffissoData, suffissoOra, risultato, saveToDB);
            
            //se l'informazione è visibile la devo scrivere anche nei fogli dei mercati
            DataView informazioni = new DataView(DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE]);
            informazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' OR SiglaEntitaRif = '" + siglaEntita + "' AND SiglaInformazione = '" + info["SiglaInformazione"] + "'";
            bool visible = false;
            foreach (DataRowView r in informazioni)
                if (r["Visibile"].Equals("1"))
                    visible = true;

            if (visible)
            {
                DataTable entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA];

                var rif =
                    (from r in entita.AsEnumerable()
                     where r["SiglaEntita"].Equals(siglaEntita)
                     select new { SiglaEntita = r["Gerarchia"] is DBNull ? r["SiglaEntita"] : r["Gerarchia"], Riferimento = r["Riferimento"] }).First();

                string quarter = Regex.Match(info["SiglaInformazione"].ToString(), @"Q\d").Value;
                quarter = quarter == "" ? "Q1" : quarter;

                Range rngMercato = new Range(_definedNamesSheetMercato.GetRowByName(rif.SiglaEntita, "UM", "T") + 2, _definedNamesSheetMercato.GetColFromName("RIF" + rif.Riferimento, "PROGRAMMA" + quarter));
                rngMercato.StartRow += (Date.GetOraFromSuffissoOra(suffissoOra) - 1);

                _wsMercato.Range[rngMercato.ToString()].Value = risultato;
            }
        }
    }
}
