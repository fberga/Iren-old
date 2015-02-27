using System;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Base
{
    public class Optimizer : CommonFunctions
    {
        private static void DeleteExistingAdjust()
        {
            DataView entitaInformazioni = LocalDB.Tables[Tab.ENTITAINFORMAZIONE].DefaultView;
            entitaInformazioni.RowFilter = "WB_Adjust <> '0'";

            foreach (DataRowView info in entitaInformazioni)
            {
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                string nomeFoglio = DefinedNames.GetSheetName(siglaEntitaInfo);
                DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);
                Tuple<int, int>[] riga = nomiDefiniti.Get(DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), info["DATA0H24"].Equals("0"));
                string strRiga = "'" + nomeFoglio + "'!" + Sheet.R1C1toA1(riga[0].Item1, riga[0].Item2) + ":" + Sheet.R1C1toA1(riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2);

                WB.Application.Run("wbAdjust", strRiga, "Reset");
                if (info["WB_Adjust"].Equals("2"))
                {
                    try
                    {
                        WB.Names.Item("WBFREE" + DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"])).Delete();
                    }
                    catch { }
                }
            }
        }
        private static void OmitConstraints()
        {
            DataView entitaInformazioni = LocalDB.Tables[Tab.ENTITAINFORMAZIONE].DefaultView;
            entitaInformazioni.RowFilter = "SiglaTipologiaInformazione = 'VINCOLO'";

            foreach (DataRowView info in entitaInformazioni)
            {
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                string nomeFoglio = DefinedNames.GetSheetName(siglaEntitaInfo);
                DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);
                Tuple<int, int>[] riga = nomiDefiniti.Get(DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), info["DATA0H24"].Equals("0"));
                string strRiga = "'" + nomeFoglio + "'!" + Sheet.R1C1toA1(riga[0].Item1, riga[0].Item2) + ":" + Sheet.R1C1toA1(riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2);

                WB.Application.Run("WBOMIT", DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), strRiga);
            }
        }
        private static void AddAdjust(object siglaEntita)
        {
            DataView entitaInformazioni = LocalDB.Tables[Tab.ENTITAINFORMAZIONE].DefaultView;
            entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND WB_Adjust <> '0'";

            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);

            foreach (DataRowView info in entitaInformazioni)
            {
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                
                Tuple<int, int>[] riga = nomiDefiniti.Get(DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), info["DATA0H24"].Equals("0"));
                string strRiga = "'" + nomeFoglio + "'!" + Sheet.R1C1toA1(riga[0].Item1, riga[0].Item2) + ":" + Sheet.R1C1toA1(riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2);

                WB.Application.Run("wbAdjust", strRiga);
                if (info["WB_Adjust"].Equals("2"))
                    WB.Application.Run("WBFREE", DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), strRiga);
            }
        }
        private static void AddConstraints(object siglaEntita)
        {
            DataView entitaInformazioni = LocalDB.Tables[Tab.ENTITAINFORMAZIONE].DefaultView;
            entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaTipologiaInformazione = 'VINCOLO'";

            foreach (DataRowView info in entitaInformazioni)
            {
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                string nomeFoglio = DefinedNames.GetSheetName(siglaEntitaInfo);
                DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);
                Tuple<int, int>[] riga = nomiDefiniti.Get(DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), info["DATA0H24"].Equals("0"));
                string strRiga = "'" + nomeFoglio + "'!" + Sheet.R1C1toA1(riga[0].Item1, riga[0].Item2) + ":" + Sheet.R1C1toA1(riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2);

                WB.Names.Item("WBOMIT" + DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"])).Delete();
            }
        }
        private static void AddOpt(object siglaEntita) 
        {
            DataView entitaInformazioni = LocalDB.Tables[Tab.ENTITAINFORMAZIONE].DefaultView;
            entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaTipologiaInformazione = 'OTTIMO'";

            if (entitaInformazioni.Count > 0)
            {
                object siglaEntitaInfo = entitaInformazioni[0]["SiglaEntitaRif"] is DBNull ? entitaInformazioni[0]["SiglaEntita"] : entitaInformazioni[0]["SiglaEntitaRif"];
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntitaInfo);
                DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);

                Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(siglaEntitaInfo, entitaInformazioni[0]["SiglaInformazione"])][0];
                string strCella = "'" + nomeFoglio + "'!" + Sheet.R1C1toA1(cella.Item1, cella.Item2);

                try { WB.Names.Item("WBMAX").Delete(); }
                catch { }

                WB.Application.Run("wbBest", strCella, "Maximize");
            }
        }
        private static void Execute(object siglaEntita)
        {
            //mantengo il filtro applicato in AddOpt
            DataView entitaInformazioni = LocalDB.Tables[Tab.ENTITAINFORMAZIONE].DefaultView;            

            if (entitaInformazioni.Count > 0)
            {
                object siglaEntitaInfo = entitaInformazioni[0]["SiglaEntitaRif"] is DBNull ? entitaInformazioni[0]["SiglaEntita"] : entitaInformazioni[0]["SiglaEntitaRif"];
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntitaInfo);
                DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);

                Excel.Worksheet ws = WB.Sheets[nomeFoglio];

                Tuple<int, int>[] riga = nomiDefiniti.Get(DefinedNames.GetName(siglaEntitaInfo, "TEMP_PROG15"), true);
                
                //eseguo con prezzi a 0
                ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Value = 1;
                WB.Application.Run("wbsolve", Arg3: "1");

                //eseguo con prezzi a 500
                ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Value = 2;
                WB.Application.Run("wbsolve", Arg3: "1");

                //eseguo con previsione prezzi
                ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Value = 3;
                WB.Application.Run("wbsolve", Arg3: "1");
            }
        }

        public static void EseguiOttimizzazione(object siglaEntita)
        {
            WB.Application.Run("wbSetGeneralOptions", Arg13: "1");

            DeleteExistingAdjust();
            OmitConstraints();
            AddAdjust(siglaEntita);
            AddConstraints(siglaEntita);
            AddOpt(siglaEntita);
            Execute(siglaEntita);

            Excel.Style style = WB.Styles["Adjustable"];
            style.Font.Bold = true;
            style.Interior.ColorIndex = 44;
        }


    }
}
