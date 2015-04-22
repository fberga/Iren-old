using System;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Iren.ToolsExcel.Utility;

namespace Iren.ToolsExcel.Base
{
    public interface IOptimizer
    {
        void EseguiOttimizzazione(object siglaEntita);
    }

    public class Optimizer : IOptimizer
    {
        DataSet _localDB;
        DataView _entitaInformazioni;
        DataView _entitaProprieta;
        string _sheet;
        NewDefinedNames _newNomiDefiniti;
        DateTime _dataFine;
        

        public Optimizer() 
        {
            _localDB = Utility.DataBase.LocalDB;
            _entitaInformazioni = _localDB.Tables[Utility.DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
            _entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
        }

        private void Helper(DataRowView info, ref string siglaEntita, ref string nomeFoglio, ref DateTime dataFine, ref NewDefinedNames newNomiDefiniti)
        {
            if (!info["SiglaEntita"].Equals(siglaEntita))
            {
                siglaEntita = info["SiglaEntita"].ToString();
                _entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                if (_entitaProprieta.Count > 0)
                    dataFine = DataBase.DataAttiva.AddDays(int.Parse(_entitaProprieta[0]["Valore"].ToString()));
                else
                    dataFine = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni);

                nomeFoglio = NewDefinedNames.GetSheetName(siglaEntita);
                newNomiDefiniti = new NewDefinedNames(nomeFoglio);
            }
        }

        protected virtual void DeleteExistingAdjust() 
        {
            _entitaInformazioni.RowFilter = "WB <> '0'";

            string siglaEntita = "";
            string nomeFoglio = "";
            DateTime dataFine = new DateTime();
            NewDefinedNames newNomiDefiniti = null;
            
            foreach (DataRowView info in _entitaInformazioni)
            {
                Helper(info, ref siglaEntita, ref nomeFoglio, ref dataFine, ref newNomiDefiniti);
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? siglaEntita : info["SiglaEntitaRif"];
                Range rng = newNomiDefiniti.Get(siglaEntitaInfo, info["SiglaInformazione"], Date.GetSuffissoDATA1).Extend(colOffset: Date.GetOreIntervallo(dataFine));
                Workbook.WB.Application.Run("wbAdjust", "'" + nomeFoglio + "'!" + rng.ToString(), "Reset");
                Workbook.WB.Sheets[nomeFoglio].Range[rng.ToString()].Style = "allDatiStyle";

                for (DateTime giorno = DataBase.DataAttiva; giorno <= dataFine; giorno = giorno.AddDays(1))
                {
                    Range rng1 = new Range(rng.StartRow, newNomiDefiniti.GetColFromDate(Date.GetSuffissoData(giorno), Date.GetSuffissoOra(Date.GetOreGiorno(giorno))));
                    Workbook.WB.Sheets[nomeFoglio].Range[rng1.ToString()].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                }

                if (info["WB"].Equals("2"))
                {
                    try
                    {
                        Workbook.WB.Names.Item("WBFREE" + NewDefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"])).Delete();
                    }
                    catch { }
                }
            }
        }

        protected virtual void OmitConstraints() 
        {
            _entitaInformazioni.RowFilter = "SiglaTipologiaInformazione = 'VINCOLO'";

            string siglaEntita = "";
            string nomeFoglio = "";
            DateTime dataFine = new DateTime();
            NewDefinedNames newNomiDefiniti = null;

            foreach (DataRowView info in _entitaInformazioni)
            {
                Helper(info, ref siglaEntita, ref nomeFoglio, ref dataFine, ref newNomiDefiniti);
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                Range rng = newNomiDefiniti.Get(siglaEntitaInfo, info["SiglaInformazione"], Date.GetSuffissoDATA1).Extend(colOffset: Date.GetOreIntervallo(dataFine));

                Workbook.WB.Application.Run("WBOMIT", NewDefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), "'" + nomeFoglio + "'!" + rng.ToString());
            }
        }
        protected virtual void AddAdjust(object siglaEntita) 
        {
            _entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND WB <> '0'";
            foreach (DataRowView info in _entitaInformazioni)
            {
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                Range rng = _newNomiDefiniti.Get(siglaEntitaInfo, info["SiglaInformazione"], Date.GetSuffissoDATA1).Extend(colOffset: Date.GetOreIntervallo(_dataFine));
                Workbook.WB.Application.Run("wbAdjust", "'" + _sheet + "'!" + rng.ToString());

                for (DateTime giorno = DataBase.DataAttiva; giorno <= _dataFine; giorno = giorno.AddDays(1))
                {
                    Range rng1 = new Range(rng.StartRow, _newNomiDefiniti.GetColFromDate(Date.GetSuffissoData(giorno), Date.GetSuffissoOra(Date.GetOreGiorno(giorno))));
                    Workbook.WB.Sheets[_sheet].Range[rng1.ToString()].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                }

                if (info["WB"].Equals("2"))
                    Workbook.WB.Application.Run("WBFREE", NewDefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), "'" + _sheet + "'!" + rng.ToString());
            }
        }
        protected virtual void AddConstraints(object siglaEntita) 
        {
            _entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaTipologiaInformazione = 'VINCOLO'";

            foreach (DataRowView info in _entitaInformazioni)
            {
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                Workbook.WB.Names.Item("WBOMIT" + NewDefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"])).Delete();
            }
        }
        protected virtual void AddOpt(object siglaEntita) 
        {
            _entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaTipologiaInformazione = 'OTTIMO'";

            if (_entitaInformazioni.Count > 0)
            {
                object siglaEntitaInfo = _entitaInformazioni[0]["SiglaEntitaRif"] is DBNull ? _entitaInformazioni[0]["SiglaEntita"] : _entitaInformazioni[0]["SiglaEntitaRif"];
                Range rng = new Range(_newNomiDefiniti.GetRowByName(siglaEntitaInfo, _entitaInformazioni[0]["SiglaInformazione"]), _newNomiDefiniti.GetFirstCol());
                try { Workbook.WB.Names.Item("WBMAX").Delete(); }
                catch { }
                Workbook.WB.Application.Run("wbBest", "'" + _sheet + "'!" + rng.ToString(), "Maximize");
            }
        }
        protected virtual void Execute(object siglaEntita) 
        {
            //mantengo il filtro applicato in AddOpt
            if (_entitaInformazioni.Count > 0)
            {
                object siglaEntitaInfo = _entitaInformazioni[0]["SiglaEntitaRif"] is DBNull ? _entitaInformazioni[0]["SiglaEntita"] : _entitaInformazioni[0]["SiglaEntitaRif"];
                Excel.Worksheet ws = Workbook.WB.Sheets[_sheet];


                if (siglaEntitaInfo.Equals("GRUPPO_TORINO"))
                {
                    Range rng = _newNomiDefiniti.Get(siglaEntitaInfo, "TEMP_PROG15", Date.GetSuffissoDATA1).Extend(colOffset: Date.GetOreIntervallo(_dataFine));

                    //eseguo con prezzi a 0
                    ws.Range[rng.ToString()].Value = 1;
                    Workbook.WB.Application.Run("wbsolve", Arg3: "1");

                    //eseguo con prezzi a 500
                    ws.Range[rng.ToString()].Value = 2;
                    Workbook.WB.Application.Run("wbsolve", Arg3: "1");

                    //eseguo con previsione prezzi
                    ws.Range[rng.ToString()].Value = 3;
                    Workbook.WB.Application.Run("wbsolve", Arg3: "1");
                }
                else
                {
                    Workbook.WB.Application.Run("wbsolve", Arg3: "1");
                }
            }
        }
        public virtual void EseguiOttimizzazione(object siglaEntita) 
        {
            Workbook.WB.Application.Run("wbSetGeneralOptions", Arg13: "1");

            _sheet = NewDefinedNames.GetSheetName(siglaEntita);
            _newNomiDefiniti = new NewDefinedNames(_sheet);

            _entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
            if (_entitaProprieta.Count > 0)
                _dataFine = DataBase.DataAttiva.AddDays(int.Parse(_entitaProprieta[0]["Valore"].ToString()));
            else
                _dataFine = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni);

            DeleteExistingAdjust();
            OmitConstraints();
            AddAdjust(siglaEntita);

            Excel.Style style = Workbook.WB.Styles["Adjustable"];

            foreach (Excel.Style s in Workbook.WB.Styles)
            {
                string sss = s.Name;
            }

            AddConstraints(siglaEntita);
            AddOpt(siglaEntita);
            Execute(siglaEntita);

            style = Workbook.WB.Styles["Adjustable"];
        }
    }
}
