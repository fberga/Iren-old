using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Classe base con i metodi per la creazione di un foglio contenente dati riferiti a impianti.
    /// </summary>
    public class Sheet : Base.Sheet
    {
        #region Costruttori

        public Sheet(Excel.Worksheet ws) 
            : base(ws)
        {
            
        }
        #endregion

        #region Metodi

        //06/02/2017 MOD: nascondo le righe dei mercati non di competenza.
        public void HideMarketRows()
        {
            /* Recupero mercato attivo al momento:
             *  - Prendo il primo mercato disponibile con chiusura > di ora
             */
            int hour = DateTime.Now.Hour;

            string mercatoAttivo = Simboli.MercatiMI
                .Where(kv => kv.Value.Chiusura > hour)
                .Select(kv => kv.Key)
                .FirstOrDefault();

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND IdApplicazione = " + Workbook.IdApplicazione; 

            foreach (DataRowView entita in categoriaEntita)
            {
                //si tratta di un'informazione di mercato (tutte le info con _MI e una cifra e visibili)
                DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaInformazione LIKE '%_MI%' AND Visibile = '1' AND IdApplicazione = " + Workbook.IdApplicazione;

                foreach (DataRowView info in informazioni)
                {
                    object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                    int row = _definedNames.GetRowByName(siglaEntita, info["SiglaInformazione"]);
                    string mercato = Regex.Match(info["SiglaInformazione"].ToString(), @"_MI\d").Value.Replace("_","");
                    _ws.Rows[row].EntireRow.Hidden = mercato != mercatoAttivo;
                }
            }
        }

        #endregion
    }
}
