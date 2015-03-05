using System;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Iren.ToolsExcel.Utility;

namespace Iren.ToolsExcel.Base
{
    public interface IOptimizer
    {
        //void DeleteExistingAdjust();
        //void OmitConstraints();
        //void AddAdjust(object siglaEntita);
        //void AddConstraints(object siglaEntita);
        //void AddOpt(object siglaEntita);
        //void Execute(object siglaEntita);
        void EseguiOttimizzazione(object siglaEntita);
    }

    public class Optimizer : IOptimizer
    {
        DataSet _localDB;
        DataView _entitaInformazioni;

        public Optimizer() 
        {
            _localDB = Utility.DataBase.LocalDB;
            _entitaInformazioni = _localDB.Tables[Utility.DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
        }

        protected virtual void DeleteExistingAdjust() 
        {
            
            _entitaInformazioni.RowFilter = "WB_Adjust <> '0'";

            foreach (DataRowView info in _entitaInformazioni)
            {
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                string nomeFoglio = DefinedNames.GetSheetName(siglaEntitaInfo);
                DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);
                Tuple<int, int>[] riga = nomiDefiniti.Get(DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), info["DATA0H24"].Equals("0"));
                string strRiga = "'" + nomeFoglio + "'!" + Sheet.R1C1toA1(riga[0].Item1, riga[0].Item2) + ":" + Sheet.R1C1toA1(riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2);

                Workbook.WB.Application.Run("wbAdjust", strRiga, "Reset");
                if (info["WB_Adjust"].Equals("2"))
                {
                    try
                    {
                        Workbook.WB.Names.Item("WBFREE" + DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"])).Delete();
                    }
                    catch { }
                }
            }
        }
        protected virtual void OmitConstraints() 
        {
            
            _entitaInformazioni.RowFilter = "SiglaTipologiaInformazione = 'VINCOLO'";

            foreach (DataRowView info in _entitaInformazioni)
            {
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                string nomeFoglio = DefinedNames.GetSheetName(siglaEntitaInfo);
                DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);
                Tuple<int, int>[] riga = nomiDefiniti.Get(DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), info["DATA0H24"].Equals("0"));
                string strRiga = "'" + nomeFoglio + "'!" + Sheet.R1C1toA1(riga[0].Item1, riga[0].Item2) + ":" + Sheet.R1C1toA1(riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2);

                Workbook.WB.Application.Run("WBOMIT", DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), strRiga);
            }
        }
        protected virtual void AddAdjust(object siglaEntita) 
        {
            
            _entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND WB_Adjust <> '0'";

            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);

            foreach (DataRowView info in _entitaInformazioni)
            {
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                
                Tuple<int, int>[] riga = nomiDefiniti.Get(DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), info["DATA0H24"].Equals("0"));
                string strRiga = "'" + nomeFoglio + "'!" + Sheet.R1C1toA1(riga[0].Item1, riga[0].Item2) + ":" + Sheet.R1C1toA1(riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2);

                Workbook.WB.Application.Run("wbAdjust", strRiga);
                if (info["WB_Adjust"].Equals("2"))
                    Workbook.WB.Application.Run("WBFREE", DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), strRiga);
            }
        }
        protected virtual void AddConstraints(object siglaEntita) 
        {
            
            _entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaTipologiaInformazione = 'VINCOLO'";

            foreach (DataRowView info in _entitaInformazioni)
            {
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                string nomeFoglio = DefinedNames.GetSheetName(siglaEntitaInfo);
                DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);
                Tuple<int, int>[] riga = nomiDefiniti.Get(DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), info["DATA0H24"].Equals("0"));
                string strRiga = "'" + nomeFoglio + "'!" + Sheet.R1C1toA1(riga[0].Item1, riga[0].Item2) + ":" + Sheet.R1C1toA1(riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2);

                Workbook.WB.Names.Item("WBOMIT" + DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"])).Delete();
            }
        }
        protected virtual void AddOpt(object siglaEntita) 
        {
            
            _entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaTipologiaInformazione = 'OTTIMO'";

            if (_entitaInformazioni.Count > 0)
            {
                object siglaEntitaInfo = _entitaInformazioni[0]["SiglaEntitaRif"] is DBNull ? _entitaInformazioni[0]["SiglaEntita"] : _entitaInformazioni[0]["SiglaEntitaRif"];
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntitaInfo);
                DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);

                Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(siglaEntitaInfo, _entitaInformazioni[0]["SiglaInformazione"])][0];
                string strCella = "'" + nomeFoglio + "'!" + Sheet.R1C1toA1(cella.Item1, cella.Item2);

                try { Workbook.WB.Names.Item("WBMAX").Delete(); }
                catch { }

                Workbook.WB.Application.Run("wbBest", strCella, "Maximize");
            }
        }
        protected virtual void Execute(object siglaEntita) 
        {
            //mantengo il filtro applicato in AddOpt
                        

            if (_entitaInformazioni.Count > 0)
            {
                object siglaEntitaInfo = _entitaInformazioni[0]["SiglaEntitaRif"] is DBNull ? _entitaInformazioni[0]["SiglaEntita"] : _entitaInformazioni[0]["SiglaEntitaRif"];
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntitaInfo);
                DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);

                Excel.Worksheet ws = Workbook.WB.Sheets[nomeFoglio];

                Tuple<int, int>[] riga = nomiDefiniti.Get(DefinedNames.GetName(siglaEntitaInfo, "TEMP_PROG15"), true);
                
                //eseguo con prezzi a 0
                ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Value = 1;
                Workbook.WB.Application.Run("wbsolve", Arg3: "1");

                //eseguo con prezzi a 500
                ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Value = 2;
                Workbook.WB.Application.Run("wbsolve", Arg3: "1");

                //eseguo con previsione prezzi
                ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Value = 3;
                Workbook.WB.Application.Run("wbsolve", Arg3: "1");
            }
        }
        public virtual void EseguiOttimizzazione(object siglaEntita) 
        {
            Workbook.WB.Application.Run("wbSetGeneralOptions", Arg13: "1");

            DeleteExistingAdjust();
            OmitConstraints();
            AddAdjust(siglaEntita);
            AddConstraints(siglaEntita);
            AddOpt(siglaEntita);
            Execute(siglaEntita);
        }
    }
}
