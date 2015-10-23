﻿using System;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Iren.ToolsExcel.Utility;
using System.Linq;

namespace Iren.ToolsExcel.Base
{
    /// <summary>
    /// Interfaccia dell'ottimizzatore.
    /// </summary>
    public interface IOptimizer
    {
        void EseguiOttimizzazione(object siglaEntita);
    }

    /// <summary>
    /// Classe per l'esecuzione dell'Ottimizzatore. All'occorrenza può essere sovrascritta da una custom inserita all'interno del pacchetto del documento.
    /// </summary>
    public class Optimizer : IOptimizer
    {
        #region Variabili

        DataView _entitaInformazioni;
        DataView _entitaProprieta;
        string _sheet;
        DefinedNames _definedNames;
        DateTime _dataFine;

        #endregion

        #region Costruttori

        public Optimizer() 
        {            
            _entitaInformazioni = Utility.DataBase.LocalDB.Tables[Utility.DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;
            _entitaProprieta = DataBase.LocalDB.Tables[DataBase.TAB.ENTITA_PROPRIETA].DefaultView;
        }

        #endregion

        #region Metodi Privati

        /// <summary>
        /// Controlla da LocalDB se ho cambiato entità dal ciclo precedente. Se sì, aggiorno la sigla entità, il nome del foglio, la data fine ed inizializzo la struttra dei nomi con questi parametri. Tutti parametri sono passati per riferimento e servono nei passaggi successivi dell'algoritmo.
        /// </summary>
        /// <param name="info">La riga delle informazioni da cui prendere i dati</param>
        /// <param name="siglaEntita">La variabile su cui salvare la Sigla Entità</param>
        /// <param name="nomeFoglio">La variabile su cui salvare il nome del foglio</param>
        /// <param name="dataFine">La variabile su cui salvare la data fine</param>
        /// <param name="definedNames">La struttura dei nomi inizializzata sul nuovo foglio</param>
        private void Helper(DataRowView info, ref string siglaEntita, ref string nomeFoglio, ref DateTime dataFine, ref DefinedNames definedNames)
        {
            if (!info["SiglaEntita"].Equals(siglaEntita))
            {
                siglaEntita = info["SiglaEntita"].ToString();
                _entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta LIKE '%GIORNI_STRUTTURA' AND IdApplicazione = " + Workbook.IdApplicazione;
                if (_entitaProprieta.Count > 0)
                    dataFine = Workbook.DataAttiva.AddDays(int.Parse(_entitaProprieta[0]["Valore"].ToString()));
                else
                    dataFine = Workbook.DataAttiva.AddDays(Struct.intervalloGiorni);

                nomeFoglio = DefinedNames.GetSheetName(siglaEntita);

                if (definedNames == null || nomeFoglio != definedNames.Sheet)
                    definedNames = new DefinedNames(nomeFoglio);
            }
        }

        #endregion

        #region Metodi Virtuali

        /// <summary>
        /// Blocca le aree su cui non considerare i vincoli.
        /// </summary>
        protected virtual void OmitConstraints() 
        {
            _entitaInformazioni.RowFilter = "SiglaTipologiaInformazione = 'VINCOLO' AND IdApplicazione = " + Workbook.IdApplicazione;

            string siglaEntita = "";
            string nomeFoglio = "";
            DateTime dataFine = new DateTime();
            DefinedNames definedNames = null;

            foreach (DataRowView info in _entitaInformazioni)
            {
                Helper(info, ref siglaEntita, ref nomeFoglio, ref dataFine, ref definedNames);
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                Range rng = definedNames.Get(siglaEntitaInfo, info["SiglaInformazione"], Date.SuffissoDATA1).Extend(colOffset: Date.GetOreIntervallo(dataFine));

                Workbook.Application.Run("WBOMIT", DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), "'" + nomeFoglio + "'!" + rng.ToString());
            }
        }
        /// <summary>
        /// Aggiunge gli adjust necessari all'entità da ottimizzare.
        /// </summary>
        /// <param name="siglaEntita">Entità da ottimizzare.</param>
        protected virtual void AddAdjust(object siglaEntita) 
        {
            _entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND WB <> '0' AND IdApplicazione = " + Workbook.IdApplicazione;
            foreach (DataRowView info in _entitaInformazioni)
            {
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                Range rng = _definedNames.Get(siglaEntitaInfo, info["SiglaInformazione"], Date.SuffissoDATA1).Extend(colOffset: Date.GetOreIntervallo(_dataFine));
                Workbook.Application.Run("wbAdjust", "'" + _sheet + "'!" + rng.ToString());

                for (DateTime giorno = Workbook.DataAttiva; giorno <= _dataFine; giorno = giorno.AddDays(1))
                {
                    Range rng1 = new Range(rng.StartRow, _definedNames.GetColFromDate(Date.GetSuffissoData(giorno), Date.GetSuffissoOra(Date.GetOreGiorno(giorno))));
                    Workbook.Sheets[_sheet].Range[rng1.ToString()].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                }

                if (info["WB"].Equals("2"))
                    Workbook.Application.Run("WBFREE", DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), "'" + _sheet + "'!" + rng.ToString());
                else if(info["WB"].Equals("3"))
                    Workbook.Application.Run("WBBIN", DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"]), "'" + _sheet + "'!" + rng.ToString());
            }
        }
        /// <summary>
        /// Aggiunge i vincoli necessari all'entità da ottimizzare.
        /// </summary>
        /// <param name="siglaEntita">Entità da ottimizzare.</param>
        protected virtual void AddConstraints(object siglaEntita) 
        {
            _entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaTipologiaInformazione = 'VINCOLO' AND IdApplicazione = " + Workbook.IdApplicazione;

            foreach (DataRowView info in _entitaInformazioni)
            {
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                Workbook.WB.Names.Item("WBOMIT" + DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"])).Delete();
            }
        }
        /// <summary>
        /// Aggiunge la funzione ottimo per l'entità da ottimizzare.
        /// </summary>
        /// <param name="siglaEntita">Entità da ottimizzare.</param>
        protected virtual void AddOpt(object siglaEntita) 
        {
            _entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaTipologiaInformazione = 'OTTIMO' AND IdApplicazione = " + Workbook.IdApplicazione;

            if (_entitaInformazioni.Count > 0)
            {
                object siglaEntitaInfo = _entitaInformazioni[0]["SiglaEntitaRif"] is DBNull ? _entitaInformazioni[0]["SiglaEntita"] : _entitaInformazioni[0]["SiglaEntitaRif"];
                Range rng = new Range(_definedNames.GetRowByName(siglaEntitaInfo, _entitaInformazioni[0]["SiglaInformazione"]), _definedNames.GetFirstCol());
                try { Workbook.WB.Names.Item("WBMAX").Delete(); }
                catch { }
                double width = Workbook.Sheets[_sheet].Range[rng.ToString()].ColumnWidth;
                Workbook.Application.Run("wbBest", "'" + _sheet + "'!" + rng.ToString(), "Maximize");
                Workbook.Sheets[_sheet].Range[rng.ToString()].ColumnWidth = width;
            }
        }
        /// <summary>
        /// Crea i MessageBox in cui avvisare l'utente dell'eventuale errore nel processo di ottimizzazione.
        /// </summary>
        /// <param name="res">Codice di errore di What's Best!.</param>
        /// <param name="messaggio">Messaggio da visualizzare.</param>
        protected virtual void ShowErrorMessageBox(int res, string messaggio)
        {
            switch (res)
            {
                case 3:
                    System.Windows.Forms.MessageBox.Show(messaggio + ": infattibile", Simboli.nomeApplicazione + " ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    break;
                case 4:
                    System.Windows.Forms.MessageBox.Show(messaggio + ": troppe soluzioni", Simboli.nomeApplicazione + " ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    break;
                case 6:
                    System.Windows.Forms.MessageBox.Show(messaggio + ": infattibile o troppe soluzioni", Simboli.nomeApplicazione + " ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    break;
            }
        }
        /// <summary>
        /// Lancia l'ottimizzazione per l'entità selezionata.
        /// </summary>
        /// <param name="siglaEntita">Entità selezionata.</param>
        protected virtual void Execute(object siglaEntita) 
        {
            //mantengo il filtro applicato in AddOpt
            if (_entitaInformazioni.Count > 0)
            {
                object siglaEntitaInfo = _entitaInformazioni[0]["SiglaEntitaRif"] is DBNull ? _entitaInformazioni[0]["SiglaEntita"] : _entitaInformazioni[0]["SiglaEntitaRif"];
                Excel.Worksheet ws = Workbook.Sheets[_sheet];


                if (siglaEntitaInfo.Equals("GRUPPO_TORINO"))
                {
                    Range rng = _definedNames.Get(siglaEntitaInfo, "TEMP_PROG15", Date.SuffissoDATA1).Extend(colOffset: Date.GetOreIntervallo(_dataFine));

                    int res = 0;

                    //eseguo con prezzi a 0
                    ws.Range[rng.ToString()].Value = 1;
                    res = Workbook.Application.Run("wbsolve", Arg3: "1");

                    ShowErrorMessageBox(res, "Calcolo dell'ottimo (prezzo 0)");

                    //eseguo con prezzi a 500
                    ws.Range[rng.ToString()].Value = 2;
                    res = Workbook.Application.Run("wbsolve", Arg3: "1");

                    ShowErrorMessageBox(res, "Calcolo dell'ottimo (prezzo 500)");

                    //eseguo con previsione prezzi
                    ws.Range[rng.ToString()].Value = 3;
                    res = Workbook.Application.Run("wbsolve", Arg3: "1");

                    ShowErrorMessageBox(res, "Calcolo dell'ottimo (previsione prezzi)");
                }
                else
                {
                    int res = Workbook.Application.Run("wbsolve", Arg3: "1");

                    ShowErrorMessageBox(res, "Calcolo dell'ottimo");
                }                
            }
        }
        /// <summary>
        /// Cancella tutti gli adjust esistenti. 
        /// </summary>
        protected virtual void DeleteExistingAdjust()
        {
            _entitaInformazioni.RowFilter = "WB <> '0' AND IdApplicazione = " + Workbook.IdApplicazione;

            string siglaEntita = "";
            string nomeFoglio = "";
            DateTime dataFine = new DateTime();
            DefinedNames definedNames = null;

            foreach (DataRowView info in _entitaInformazioni)
            {
                Helper(info, ref siglaEntita, ref nomeFoglio, ref dataFine, ref definedNames);
                object siglaEntitaInfo = info["SiglaEntitaRif"] is DBNull ? siglaEntita : info["SiglaEntitaRif"];
                Range rng = definedNames.Get(siglaEntitaInfo, info["SiglaInformazione"], Date.SuffissoDATA1).Extend(colOffset: Date.GetOreIntervallo(dataFine));
                double width = Workbook.Sheets[nomeFoglio].Range[rng.ToString()].ColumnWidth;
                Workbook.Application.Run("wbAdjust", "'" + nomeFoglio + "'!" + rng.ToString(), "Reset");
                Workbook.Sheets[nomeFoglio].Range[rng.ToString()].ColumnWidth = width;
                Workbook.Sheets[nomeFoglio].Range[rng.ToString()].Style = "Area dati";

                for (DateTime giorno = Workbook.DataAttiva; giorno <= dataFine; giorno = giorno.AddDays(1))
                {
                    Range rng1 = new Range(rng.StartRow, definedNames.GetColFromDate(Date.GetSuffissoData(giorno), Date.GetSuffissoOra(Date.GetOreGiorno(giorno))));
                    Workbook.Sheets[nomeFoglio].Range[rng1.ToString()].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                }

                if (info["WB"].Equals("2"))
                {
                    try
                    {
                        Workbook.WB.Names.Item("WBFREE" + DefinedNames.GetName(siglaEntitaInfo, info["SiglaInformazione"])).Delete();
                    }
                    catch { }
                }
            }
        }
        /// <summary>
        /// Funizione ereditata dall'interfaccia che viene richiamata nella parte base dell'algoritmo per eseguire l'ottimizzazione.
        /// </summary>
        /// <param name="siglaEntita">Entità da ottimizzare.</param>
        public virtual void EseguiOttimizzazione(object siglaEntita) 
        {
            Workbook.Application.Run("wbSetGeneralOptions", Arg3: "120", Arg13: "1");

            _sheet = DefinedNames.GetSheetName(siglaEntita);
            _definedNames = new DefinedNames(_sheet);

            string desEntita =
                (from r in DataBase.LocalDB.Tables[DataBase.TAB.CATEGORIA_ENTITA].AsEnumerable()
                 where r["IdApplicazione"].Equals(Workbook.IdApplicazione) && r["SiglaEntita"].Equals(siglaEntita)
                 select r["DesEntita"].ToString()).First();

            _entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta LIKE '%GIORNI_STRUTTURA' AND IdApplicazione = " + Workbook.IdApplicazione;
            if (_entitaProprieta.Count > 0)
                _dataFine = Workbook.DataAttiva.AddDays(int.Parse(_entitaProprieta[0]["Valore"].ToString()));
            else
                _dataFine = Workbook.DataAttiva.AddDays(Struct.intervalloGiorni);

            OmitConstraints();
            AddAdjust(siglaEntita);
            AddConstraints(siglaEntita);
            AddOpt(siglaEntita);
            SplashScreen.Close();
            Execute(siglaEntita);
            DeleteExistingAdjust();

            Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogGenera, "Eseguita ottimizzazione " + desEntita);
        }

        #endregion
    }
}
