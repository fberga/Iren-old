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

        public override void CaricaInformazioni(bool all)
        {
            try
            {
                if (DataBase.OpenConnection())
                {
                    DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
                    DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
                    categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "'";

                    _dataInizio = DataBase.DB.DataAttiva;

                    DateTime dataFineMax = _dataInizio;
                    Dictionary<object, DateTime> dateFineUP = new Dictionary<object, DateTime>();
                    foreach (DataRowView entita in categoriaEntita)
                    {
                        entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                        if (entitaProprieta.Count > 0)
                            dateFineUP.Add(entita["SiglaEntita"], _dataInizio.AddDays(double.Parse("" + entitaProprieta[0]["Valore"])));
                        else
                            dateFineUP.Add(entita["SiglaEntita"], _dataInizio.AddDays(Struct.intervalloGiorni));

                        dataFineMax = new DateTime(Math.Max(dataFineMax.Ticks, dateFineUP[entita["SiglaEntita"]].Ticks));
                    }

                    DataView datiApplicazioneH = DataBase.Select(DataBase.SP.APPLICAZIONE_INFORMAZIONE_H, "@SiglaCategoria=" + _siglaCategoria + ";@SiglaEntita=ALL;@DateFrom=" + _dataInizio.ToString("yyyyMMdd") + ";@DateTo=" + dataFineMax.ToString("yyyyMMdd") + ";@Tipo=1;@All=" + (all ? "1" : "0")).DefaultView;

                    DataView insertManuali = new DataView();
                    if (all)
                        insertManuali = DataBase.Select(DataBase.SP.APPLICAZIONE_INFORMAZIONE_COMMENTO, "@SiglaCategoria=" + _siglaCategoria + ";@SiglaEntita=ALL;@DateFrom=" + _dataInizio.ToString("yyyyMMdd") + ";@DateTo=" + dataFineMax.ToString("yyyyMMdd") + ";@All=1").DefaultView;

                    if (Struct.tipoVisualizzazione == "O")
                    {
                        foreach (DataRowView entita in categoriaEntita)
                        {
                            SplashScreen.UpdateStatus("Carico informazioni " + entita["DesEntita"]);

                            datiApplicazioneH.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND CONVERT(Data, System.Int32) <= " + dateFineUP[entita["SiglaEntita"]].ToString("yyyyMMdd");

                            SplashScreen.UpdateStatus("Carico informazioni " + entita["SiglaEntita"]);
                            CaricaInformazioniEntita(datiApplicazioneH);
                            if (all)
                            {
                                insertManuali.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND CONVERT(SUBSTRING(Data, 1, 8), System.Int32) <= " + dateFineUP[entita["SiglaEntita"]].ToString("yyyyMMdd");
                                CaricaCommentiEntita(insertManuali);
                            }
                        }
                    }
                    else
                    {
                        SplashScreen.UpdateStatus("Carico informazioni " + _siglaCategoria);
                        CaricaInformazioniEntita(datiApplicazioneH);
                        if (all)
                        {
                            CaricaCommentiEntita(insertManuali);
                        }
                    }

                    //carico dati giornalieri
                    DataView datiApplicazioneD = DataBase.Select(DataBase.SP.APPLICAZIONE_INFORMAZIONE_D, "@SiglaCategoria=" + _siglaCategoria + ";@SiglaEntita=ALL;@DateFrom=" + DataBase.DataAttiva.ToString("yyyyMMdd") + ";@DateTo=" + DataBase.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("yyyyMMdd") + ";@Tipo=1;@All=" + (all ? "1" : "0")).DefaultView;

                    foreach (DataRowView dato in datiApplicazioneD)
                    {
                        Range rng = new Range(_definedNames.GetRowByNameSuffissoData(dato["SiglaEntita"], dato["SiglaInformazione"], Date.GetSuffissoData(dato["Data"].ToString())), _definedNames.GetFirstCol() - 1);

                        _ws.Range[rng.ToString()].Value = dato["Valore"];
                    }
                }
            }
            catch (Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "CaricaInformazioni [all = " + all + "]: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
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
                    o.RemoveRange(o.Count - 4, 4);
                    _ws.Range[rng.ToString()].Value = o.ToArray();

                    //trovo il range nel foglio nascosto
                    string entitaPadre = _definedNamesMercatoAttivo.GetFullNameByParts(dato["SiglaEntita"]).First(s => s.Contains("RIF"));
                    string[] parts = entitaPadre.Split(Simboli.UNION[0]);
                    string info = dato["SiglaInformazione"].ToString().Split('_')[0];


                    if (!Regex.IsMatch(info, @"Q\d"))
                        info += "Q1";

                    Range rngMercatoAttivo = new Range(_definedNamesMercatoAttivo.GetRowByName(parts[1], "UM", "T") + 2, _definedNamesMercatoAttivo.GetColFromName(parts[2], info), Date.GetOreGiorno(giorno));

                    //copio i dati nel foglio nascosto
                    for (int i = 0; i < Date.GetOreGiorno(giorno); i++)
                        Workbook.WB.Sheets[Simboli.Mercato].Range[rngMercatoAttivo.Rows[i].ToString()].Value = o[i];

                    if (giorno == DataBase.DataAttiva && Regex.IsMatch(dato["SiglaInformazione"].ToString(), @"RIF\d+"))
                    {
                        Selection s = _definedNames.GetSelectionByRif(rng);
                        s.ClearSelections(_ws);
                        s.Select(_ws, int.Parse(o[0].ToString().Split('.')[0]));
                    }
                }
            }
        }

        //public override void UpdateData(bool all = true)
        //{
        //    SplashScreen.UpdateStatus("Aggiorno informazioni");
        //    if (all)
        //    {
        //        CancellaDati();
        //        AggiornaDateTitoli();
        //        CaricaParametri();
        //    }
        //    CaricaInformazioni(all);
        //    AggiornaGrafici();
        //}
        //#region UpdateData

        //private void CancellaDati()
        //{
        //    CancellaDati(DataBase.DataAttiva, true);
        //}
        //private void CancellaDati(DateTime giorno, bool all = false)
        //{
        //    DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
        //    categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "'"; // AND (Gerarchia = '' OR Gerarchia IS NULL )";

        //    string suffissoData = Date.GetSuffissoData(giorno);
        //    int colOffset = _definedNames.GetColOffset();
        //    if (!all)
        //        colOffset = Date.GetOreGiorno(giorno);

        //    foreach (DataRowView entita in categoriaEntita)
        //    {
        //        SplashScreen.UpdateStatus("Cancello dati " + entita["DesEntita"]);
        //        DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
        //        informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND FormulaInCella = '0'";// AND ValoreDefault IS NULL";

        //        foreach (DataRowView info in informazioni)
        //        {
        //            int col = all ? _definedNames.GetFirstCol() : _definedNames.GetColFromDate(suffissoData);
        //            object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

        //            if (Struct.tipoVisualizzazione == "O")
        //            {
        //                int row = _definedNames.GetRowByName(siglaEntita, info["SiglaInformazione"]);
        //                if (info["SiglaTipologiaInformazione"].Equals("GIORNALIERA"))
        //                {
        //                    Excel.Range rngData = _ws.Range[Range.GetRange(row, col - 1)];
        //                    rngData.Value = "";
        //                }
        //                else
        //                {
        //                    Excel.Range rngData = _ws.Range[Range.GetRange(row, col, 1, colOffset)];
        //                    rngData.Value = "";
        //                    rngData.ClearComments();
        //                    Style.RangeStyle(rngData, backColor: info["BackColor"], foreColor: info["ForeColor"]);
        //                }
        //            }
        //            else
        //            {
        //                DateTime dataInizio = giorno;
        //                DateTime dataFine = giorno;
        //                if (all)
        //                {
        //                    dataInizio = DataBase.DataAttiva;
        //                    dataFine = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni);
        //                }

        //                CicloGiorni(dataInizio, dataFine, (oreGiorno, suffData, g) =>
        //                {
        //                    SplashScreen.UpdateStatus("Cancello dati " + g.ToShortDateString());

        //                    int row = _definedNames.GetRowByNameSuffissoData(siglaEntita, info["SiglaInformazione"], suffData);
        //                    if (info["SiglaTipologiaInformazione"].Equals("GIORNALIERA"))
        //                    {
        //                        Excel.Range rngData = _ws.Range[Range.GetRange(row, col - 1)];
        //                        rngData.Value = "";
        //                    }
        //                    else
        //                    {
        //                        Excel.Range rng = _ws.Range[Range.GetRange(row, col, 1, oreGiorno)];
        //                        rng.Value = "";
        //                        rng.ClearComments();
        //                        Style.RangeStyle(rng, backColor: info["BackColor"], foreColor: info["ForeColor"]);
        //                    }
        //                });
        //            }
        //        }
        //        //reset colonna 24esima 25esima ora
        //        if (all && Struct.tipoVisualizzazione == "V" && informazioni.Count > 0)
        //        {
        //            DateTime dataInizio = DataBase.DataAttiva;
        //            DateTime dataFine = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni);

        //            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];

        //            CicloGiorni(dataInizio, dataFine, (oreGiorno, suffData, g) =>
        //            {
        //                Range rngData = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], suffData), _definedNames.GetFirstCol(), informazioni.Count, oreGiorno);

        //                int ore = Date.GetOreGiorno(g);
        //                if (ore == 23)
        //                {
        //                    _ws.Range[rngData.Columns[rngData.Columns.Count - 2, rngData.Columns.Count - 1].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
        //                }
        //                else if (ore == 24)
        //                {
        //                    _ws.Range[rngData.Columns[rngData.Columns.Count - 2].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternNone;
        //                    _ws.Range[rngData.Columns[rngData.Columns.Count - 1].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
        //                }
        //                else if (ore == 25)
        //                {
        //                    _ws.Range[rngData.Columns[rngData.Columns.Count - 2].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternNone;
        //                    _ws.Range[rngData.Columns[rngData.Columns.Count - 1].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternNone;
        //                }
        //            });
        //        }
        //    }
        //}

        //#endregion
        
        #endregion
    }
}
