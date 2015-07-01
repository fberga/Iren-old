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
    public class Modifica : Base.Modifica
    {
        public override void Range(object Sh, Excel.Range Target)
        {
            //Se la funzione scrive in altre celle, ricordarsi di disabilitare gli handler per la modifica delle celle
            Workbook.WB.SheetChange -= Handler.StoreEdit;
            Workbook.WB.SheetChange -= this.Range;


            Excel.Worksheet ws = Target.Worksheet;
            Excel.Worksheet wsMercato = Workbook.Sheets[Simboli.Mercato];

            bool wasProtected = wsMercato.ProtectContents;
            if (wasProtected)
            {
                wsMercato.Unprotect(Simboli.pwd);
                ws.Unprotect(Simboli.pwd);
            }

            DefinedNames definedNames = new DefinedNames(ws.Name);
            DefinedNames definedNamesMercato = new DefinedNames(Simboli.Mercato);

            DataTable entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA];

            string[] ranges = Target.Address.Split(',');

            foreach (string range in ranges)
            {
                Range rng = new Range(range);

                foreach (Range cell in rng.Cells)
                {
                    string[] parts = definedNames.GetNameByAddress(cell.StartRow, cell.StartColumn).Split(Simboli.UNION[0]);

                    string siglaEntita = parts[0];
                    string siglaInformazione = parts[1];
                    string suffissoData = parts[2];
                    string suffissoOra = parts[3];

                    var rif =
                    (from r in entita.AsEnumerable()
                     where r["SiglaEntita"].Equals(siglaEntita)
                     select new { SiglaEntita = r["Gerarchia"] is DBNull ? r["SiglaEntita"] : r["Gerarchia"], Riferimento = r["Riferimento"] }).First();

                    string quarter = Regex.Match(siglaInformazione, @"Q\d").Value;
                    quarter = quarter == "" ? "Q1" : quarter;

                    Range rngMercato = new Range(definedNamesMercato.GetRowByName(rif.SiglaEntita, "UM", "T") + 2, definedNamesMercato.GetColFromName("RIF" + rif.Riferimento, "PROGRAMMA" + quarter));
                    rngMercato.StartRow += (Date.GetOraFromSuffissoOra(suffissoOra) - 1);

                    wsMercato.Range[rngMercato.ToString()].Value = ws.Range[cell.ToString()].Value;
                    ws.Range[cell.ToString()].Interior.ColorIndex = wsMercato.Range[rngMercato.ToString()].DisplayFormat.Interior.ColorIndex;
                }
            }

            if (wasProtected)
            {
                wsMercato.Protect(Simboli.pwd);
                ws.Protect(Simboli.pwd);
            }

            //Se la funzione scrive in altre celle, ricordarsi di riabilitare gli handler per la modifica delle celle
            Workbook.WB.SheetChange += Handler.StoreEdit;
            Workbook.WB.SheetChange += this.Range;
        }
    }
}
