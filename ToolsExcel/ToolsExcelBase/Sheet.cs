using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Base
{
    public abstract class ASheet
    {
        #region Variabili

        protected Struct _struttura;

        protected DateTime _dataInizio;
        protected DateTime _dataFine;
        protected int _visParametro;

        #endregion

        protected void CicloGiorni(DateTime dataInizio, DateTime dataFine, Action<int, string, DateTime> callback)
        {
            for (DateTime giorno = dataInizio; giorno <= dataFine; giorno = giorno.AddDays(1))
            {
                int oreGiorno = Date.GetOreGiorno(giorno);
                string suffissoData = Date.GetSuffissoData(_dataInizio, giorno);

                if (Struct.tipoVisualizzazione == "V")
                {
                    oreGiorno = 25;
                    suffissoData = Date.GetSuffissoData(DataBase.DataAttiva, giorno);
                }

                //TODO rimuovere
                if (giorno == _dataInizio && _struttura.visData0H24)
                {
                    oreGiorno++;
                }

                callback(oreGiorno, suffissoData, giorno);
            }
        }
        protected void CicloGiorni(Action<int, string, DateTime> callback)
        {
            CicloGiorni(_dataInizio, _dataFine, callback);
        }

        public abstract void LoadStructure();
        public abstract void UpdateData(bool all = true);
        public abstract void CalcolaFormule(string siglaEntita = null, DateTime? giorno = null, int ordineElaborazione = 0, bool escludiOrdine = false);
        public abstract void AggiornaDateTitoli();
        public abstract void AggiornaGrafici();
        protected abstract void InsertPersonalizzazioni();

        public abstract void CaricaInformazioni(bool all);

        #region Metodi Statici

        public static void Proteggi(bool proteggi)
        {
            foreach (Excel.Worksheet ws in Workbook.WB.Sheets)
            {
                if (proteggi)
                    if (ws.Name == "Log")
                        ws.Protect(Simboli.pwd, AllowSorting: true, AllowFiltering: true);
                    else
                        ws.Protect(Simboli.pwd);
                else
                    ws.Unprotect(Simboli.pwd);
            }
        }
        public static void AbilitaModifica(bool abilita)
        {
            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            categorie.RowFilter = "Operativa = '1'";

            Proteggi(false);
            foreach (DataRowView categoria in categorie)
            {
                DefinedNames nomiDefiniti = new DefinedNames(categoria["DesCategoria"].ToString());
                Excel.Worksheet ws = Workbook.WB.Sheets[categoria["DesCategoria"].ToString()];

                DataView informazioni = nomiDefiniti.GetEditable();
                foreach (DataRowView info in informazioni)
                {
                    //se i giorni sono in verticale, devo disabilitare dove necessario l'ora 24 e la 25
                    List<string> exclude = new List<string>();
                    if (Struct.tipoVisualizzazione == "V")
                    {
                        int oreGiorno = Date.GetOreGiorno(Date.GetDataFromSuffisso(info["SuffissoData"]));
                        if (oreGiorno == 23)
                        {
                            exclude.Add("H24");
                            exclude.Add("H25");
                        }
                        else if (oreGiorno == 24)
                            exclude.Add("H25");
                    }
                    Tuple<int, int>[] riga = nomiDefiniti.Get(DefinedNames.GetName(info["SiglaEntita"], info["SiglaInformazione"], info["SuffissoData"]), exclude.ToArray());
                    ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Locked = !abilita;
                }
            }
            Proteggi(true);

            //DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            //categorie.RowFilter = "Operativa = '1'";
            //DataView entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            //DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;

            //foreach (DataRowView categoria in categorie)
            //{
            //    DefinedNames nomiDefiniti = new DefinedNames(categoria["DesCategoria"].ToString());
            //    Excel.Worksheet ws = Workbook.WB.Sheets[categoria["DesCategoria"].ToString()];

            //    Proteggi(false);
            //    entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "'";
            //    foreach (DataRowView e in entita)
            //    {
            //        informazioni.RowFilter = "SiglaEntita = '" + e["SiglaEntita"] + "' AND Editabile = '1'";
            //        foreach (DataRowView info in informazioni)
            //        {
            //            object siglaEntita = info["SiglaEntitaRif"] is DBNull ? e["SiglaEntita"] : info["SiglaEntitaRif"];
            //            Tuple<int, int>[] riga = nomiDefiniti[DefinedNames.GetName(siglaEntita, info["SiglaInformazione"])];

            //            ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Locked = !abilita;
            //        }
            //    }
            //    Proteggi(true);
            //}
        }

        public static void SalvaModifiche(DateTime inizio, DateTime fine)
        {
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            DataView entitaInformazione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;

            foreach (Excel.Worksheet ws in Workbook.WB.Sheets)
            {
                if (ws.Name != "Main" && ws.Name != "Log")
                {
                    DefinedNames nomiDefiniti = new DefinedNames(ws.Name);
                    categorie.RowFilter = "DesCategoria = '" + ws.Name + "' AND Operativa = '1'";
                    categoriaEntita.RowFilter = "SiglaCategoria = '" + categorie[0]["SiglaCategoria"] + "'";

                    for (DateTime giorno = inizio; giorno <= fine; giorno = giorno.AddDays(1))
                    {
                        foreach (DataRowView entita in categoriaEntita)
                        {
                            entitaInformazione.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND FormulaInCella = '1' AND WB = '0' AND SalvaDB = '1'";
                            foreach (DataRowView info in entitaInformazione)
                            {
                                object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                                //Tuple<int, int>[] rngInfo = nomiDefiniti[DefinedNames.GetName(siglaEntita, info["SiglaInformazione"], Date.GetSuffissoData(giorno))];
                                //Excel.Range rng = ws.Range[ws.Cells[rngInfo[0].Item1, rngInfo[0].Item2], ws.Cells[rngInfo[rngInfo.Length - 1].Item1, rngInfo[rngInfo.Length - 1].Item2]];
                                //Handler.StoreEdit(ws, rng);

                                Tuple<int,int>[] rngInfo = nomiDefiniti[siglaEntita, info["SiglaInformazione"], Date.GetSuffissoData(giorno)];
                                //TODO SISTEMARE
                                //Handler.StoreEdit(ws, ws.Range[nomiDefiniti.GetRange(rngInfo)]);
                            }
                        }
                    }
                }
            }
        }
        public static void SalvaModifiche()
        {
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            DataView entitaInformazione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;

            foreach (Excel.Worksheet ws in Workbook.WB.Sheets)
            {
                if (ws.Name != "Main" && ws.Name != "Log")
                {
                    DefinedNames nomiDefiniti = new DefinedNames(ws.Name);
                    categorie.RowFilter = "DesCategoria = '" + ws.Name + "' AND Operativa = '1'";
                    categoriaEntita.RowFilter = "SiglaCategoria = '" + categorie[0]["SiglaCategoria"] + "'";

                    foreach (DataRowView entita in categoriaEntita)
                    {
                        entitaInformazione.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND FormulaInCella = '1' AND WB = '0' AND SalvaDB = '1'";
                        foreach (DataRowView info in entitaInformazione)
                        {
                            object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                            Tuple<int, int>[] rngInfo = nomiDefiniti[info["Data0H24"].Equals("0"), siglaEntita, info["SiglaInformazione"]];
                            //TODO SISTEMARE
                            //Handler.StoreEdit(ws, ws.Range[nomiDefiniti.GetRange(rngInfo)]);
                        }
                    }
                }
            }
        }

        #endregion
    }

    public class Sheet : ASheet, IDisposable
    {
        #region Variabili

        protected Excel.Worksheet _ws;
        protected DefinedNames _nomiDefiniti;
        protected NewDefinedNames _newNomiDefiniti;
        protected object _siglaCategoria;
        protected int _colonnaInizio;
        protected int _intervalloOre;
        protected int _rigaAttiva;
        protected bool _disposed = false;
        

        #endregion

        #region Costruttori

        public Sheet(Excel.Worksheet ws)
        {
            _ws = ws;

            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            categorie.RowFilter = "DesCategoria = '" + ws.Name + "'";

            _siglaCategoria = categorie[0]["SiglaCategoria"];            

            //dimensionamento celle in base ai parametri del DB
            DataView paramApplicazione = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE].DefaultView;

            _struttura = new Struct();

            _struttura.rigaBlock = (int)paramApplicazione[0]["RowBlocco"] + (paramApplicazione[0]["TipoVisualizzazione"].Equals("O") ? 2 : 0);
            _struttura.rigaGoto = (int)paramApplicazione[0]["RowGoto"];
            _struttura.visData0H24 = paramApplicazione[0]["VisData0H24"].ToString() == "1";
            _struttura.visParametro = paramApplicazione[0]["VisParametro"].ToString() == "1";
            _struttura.colBlock = (int)paramApplicazione[0]["ColBlocco"] + (_struttura.visParametro ? 1 : 0);
            Struct.tipoVisualizzazione = paramApplicazione[0]["TipoVisualizzazione"] is DBNull ? "O" : paramApplicazione[0]["TipoVisualizzazione"].ToString();
            Struct.intervalloGiorni = paramApplicazione[0]["IntervalloGiorniEntita"] is DBNull ? 0 : (int)paramApplicazione[0]["IntervalloGiorniEntita"];

            _visParametro = _struttura.visParametro ? 3 : 2;
            _nomiDefiniti = new DefinedNames(_ws.Name);
            _newNomiDefiniti = new NewDefinedNames(_ws.Name);            
        }
        ~Sheet()
        {
            Dispose();
        }

        #endregion

        #region Metodi

        public override void LoadStructure()
        {
            DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPROPRIETA].DefaultView;
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;

            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";
            _dataInizio = Utility.DataBase.DB.DataAttiva;

            //carico la massima datafine in maniera da creare la barra navigazione della dimensione giusta (compresa la definizione dei giorni se necessario)
            int intervalloGiorniMax = 0;
            if (Struct.tipoVisualizzazione == "O")
            {
                foreach (DataRowView entita in categoriaEntita)
                {
                    string siglaEntita = "" + entita["SiglaEntita"];
                    entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                    if (entitaProprieta.Count > 0)
                    {
                        intervalloGiorniMax = Math.Max(intervalloGiorniMax, int.Parse("" + entitaProprieta[0]["Valore"]));

                    }
                }
            }
            _dataFine = Utility.DataBase.DB.DataAttiva.AddDays(intervalloGiorniMax);
            _newNomiDefiniti.DefineDates(_dataInizio, _dataFine, _struttura.colBlock, _struttura.visData0H24);

            Clear();
            InitBarraNavigazione();

            _rigaAttiva = _struttura.rigaBlock + 1;

            foreach (DataRowView entita in categoriaEntita)
            {
                string siglaEntita = "" + entita["SiglaEntita"];
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";

                if (Struct.tipoVisualizzazione == "O")
                {
                    if (entitaProprieta.Count > 0)
                        _dataFine = _dataInizio.AddDays(double.Parse("" + entitaProprieta[0]["Valore"]));
                    else
                        _dataFine = _dataInizio.AddDays(Struct.intervalloGiorni);

                    InitBloccoEntita(entita);
                }
                else if (Struct.tipoVisualizzazione == "V")
                {
                    CicloGiorni(_dataInizio, _dataInizio.AddDays(Struct.intervalloGiorni), (oreGiorno, suffissoData, giorno) =>
                    {
                        _dataFine = _dataInizio = giorno;
                        InitBloccoEntita(entita);
                    });
                }
            }

            entitaProprieta.RowFilter = "";
            categoriaEntita.RowFilter = "";

            CaricaInformazioni(all: true);
            AggiornaGrafici();            
            //CalcolaFormule();                     //TODO

            _newNomiDefiniti.DumpToDataSet();
        }

        protected void Clear()
        {
            _ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            if (_ws.ChartObjects().Count > 0)
                _ws.ChartObjects().Delete();

            _ws.UsedRange.EntireColumn.Delete();
            _ws.UsedRange.FormatConditions.Delete();
            _ws.UsedRange.EntireRow.Hidden = false;
            _ws.UsedRange.Font.Size = 10;
            _ws.UsedRange.NumberFormat = "General";
            _ws.UsedRange.Font.Name = "Verdana";
            _ws.UsedRange.RowHeight = Struct.cell.height.normal;

            _ws.Columns.ColumnWidth = Struct.cell.width.dato;

            _ws.Rows["1:" + (_struttura.rigaBlock - 1)].RowHeight = Struct.cell.height.empty;
            _ws.Rows[_struttura.rigaGoto].RowHeight = Struct.cell.height.normal;

            _ws.Columns[1].ColumnWidth = Struct.cell.width.empty;
            _ws.Columns[2].ColumnWidth = Struct.cell.width.entita;

            ((Excel._Worksheet)_ws).Activate();
            _ws.Application.ActiveWindow.FreezePanes = false;
            _ws.Cells[_struttura.rigaBlock, _struttura.colBlock].Select();
            _ws.Application.ActiveWindow.ScrollColumn = 1;
            _ws.Application.ActiveWindow.ScrollRow = 1;
            _ws.Application.ActiveWindow.FreezePanes = true;

            int colInfo = _struttura.colBlock - _visParametro;
            _ws.Columns[colInfo].ColumnWidth = Struct.cell.width.informazione;
            _ws.Columns[colInfo + 1].ColumnWidth = Struct.cell.width.unitaMisura;
            if (_struttura.visParametro)
                _ws.Columns[colInfo + 2].ColumnWidth = Struct.cell.width.parametro;
        }
        protected void InitBarraNavigazione()
        {
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";

            int dataOreTot = (Struct.tipoVisualizzazione == "O" ? Date.GetOreIntervallo(_dataInizio, _dataFine) : 25) + (_struttura.visData0H24 ? 1 : 0) + (_struttura.visParametro ? 1 : 0);
            int numElementiMenu = (Struct.tipoVisualizzazione == "O" ? categoriaEntita.Count : (Struct.intervalloGiorni + 1));
                
            Excel.Range gotoBar = _ws.Range[_ws.Cells[2, 2], _ws.Cells[_struttura.rigaGoto + 1, _struttura.colBlock + dataOreTot - 1]];
            gotoBar.Style = "gotoBarStyle";
            gotoBar.BorderAround2(Weight: Excel.XlBorderWeight.xlMedium, Color: 1);

            //vedo se e come dividere gli elementi per riga
            int numRighe = 1;
            if (numElementiMenu > 8)
            {
                int tmp = numElementiMenu;
                while (tmp / 8 > 0)
                {
                    _ws.Rows[_struttura.rigaGoto + 1].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    _struttura.rigaBlock++;
                    numRighe++;
                    tmp /= 8;
                }
            }
            double numEleRiga = numElementiMenu / Convert.ToDouble(numRighe);

            int j = 0;
            for (int i = 0; i < numElementiMenu; i++)
            {
                int r = (i / (int)Math.Ceiling(numEleRiga));
                int c = (i % (int)Math.Ceiling(numEleRiga));

                object nome = Struct.tipoVisualizzazione == "O" ? categoriaEntita[i]["SiglaEntita"] : Date.GetSuffissoData(DataBase.DataAttiva.AddDays(i));
                _newNomiDefiniti.AddGOTO(nome, _struttura.rigaGoto + r, _struttura.colBlock + c);

                Excel.Range rng;
                if (Struct.cell.width.dato < 10)
                {
                    j = c == 0 ? 0 : j + 1;
                    c += j;
                    rng = _ws.Range[_ws.Cells[_struttura.rigaGoto + r, _struttura.colBlock + c + (_struttura.visData0H24 ? 1 : 0)], _ws.Cells[_struttura.rigaGoto + r, _struttura.colBlock + c + 1 + (_struttura.visData0H24 ? 1 : 0)]];
                    rng.Merge();
                }
                else
                {
                    rng = _ws.Cells[_struttura.rigaGoto + r, _struttura.colBlock + c + (_struttura.visData0H24 ? 1 : 0)];
                }
                rng.Value = Struct.tipoVisualizzazione == "O" ? categoriaEntita[i]["DesEntitaBreve"] : DataBase.DataAttiva.AddDays(i);
                rng.Style = Struct.tipoVisualizzazione == "O" ? "navBarStyleHorizontal" : "navBarStyleVertical";
            }

            //inserisco la data e le ore
            if (Struct.tipoVisualizzazione == "O")
            {
                int colonnaInizio = _struttura.colBlock;
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    int escludiH24 = (giorno == _dataInizio && _struttura.visData0H24 ? 1 : 0);

                    Range rngData = new Range(_struttura.rigaBlock - 2, colonnaInizio + escludiH24, 0, oreGiorno - 1 - escludiH24);

                    Excel.Range rng = _ws.Range[rngData.ToString()];
                    rng.Merge();
                    rng.Style = "dateBarStyle";
                    rng.Value = giorno.ToString("MM/dd/yyyy");
                    rng.RowHeight = 25;

                    Range rngOre = new Range(_struttura.rigaBlock - 1, colonnaInizio, 0, oreGiorno - 1);
                    InsertOre(rngOre, giorno == _dataInizio && _struttura.visData0H24);
                    colonnaInizio += oreGiorno;
                });
            }
        }
        
        protected void InitBloccoEntita(DataRowView entita)
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            DataView grafici = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAGRAFICO].DefaultView;
            DataView graficiInfo = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAGRAFICOINFORMAZIONE].DefaultView;

            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
            informazioni.Sort = "Ordine";

            grafici.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
            graficiInfo.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

            _colonnaInizio = _struttura.colBlock;
            _intervalloOre = Date.GetOreIntervallo(_dataInizio, _dataFine) + (_struttura.visData0H24 ? 1 : 0) + (_struttura.visParametro ? 1 : 0);

            CreaNomiCelle(entita["SiglaEntita"]);
            InsertTitoloEntita(entita["SiglaEntita"], entita["DesEntita"]);
            InsertOre(entita["SiglaEntita"]);
            InsertTitoloVerticale(entita["DesEntitaBreve"]);
            FormattaBloccoEntita();
            InsertInformazioniEntita();
            InsertPersonalizzazioni();
            InsertGrafici();
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND (ValoreDefault IS NOT NULL OR FormulaInCella = 1)";
            InsertFormuleValoriDefault();
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaTipologiaParametro IS NOT NULL";
            InsertParametri();
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
            FormattazioneCondizionale();

            //due righe vuote tra un'entità e la successiva
            _rigaAttiva += 2;
        }
        #region Blocco entità

        protected virtual void CreaNomiCelle(object siglaEntita)
        {
            //inserisco titoli
            string suffissoData = Date.GetSuffissoData(_dataInizio);
            _newNomiDefiniti.AddName(_rigaAttiva, siglaEntita, "T", Struct.tipoVisualizzazione == "O" ? "" : suffissoData);

            //sistema l'indirizzamento dei GOTO
            int col = _newNomiDefiniti.GetColFromDate(suffissoData);
            object name = Struct.tipoVisualizzazione == "O" ? siglaEntita : suffissoData;
            _newNomiDefiniti.ChangeGOTOAddressTo(name, Range.R1C1toA1(_rigaAttiva, col));

            //aggiungo la riga delle ore
            _rigaAttiva += Struct.tipoVisualizzazione == "V" ? 2 : 1;

            //aggiungo i grafici
            DataView grafici = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAGRAFICO].DefaultView;

            int i = 1;
            foreach (DataRowView grafico in grafici)
            {
                _newNomiDefiniti.AddName(_rigaAttiva, grafico["SiglaEntita"], "GRAFICO" + i, Struct.tipoVisualizzazione == "O" ? "" : Date.GetSuffissoData(_dataInizio));
                i++;
                _rigaAttiva++;
            }

            //aggiungo informazioni
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            _newNomiDefiniti.AddName(_rigaAttiva, Struct.tipoVisualizzazione == "O" ? siglaEntita : suffissoData, "TITOLO_VERTICALE");
            foreach (DataRowView info in informazioni)
            {
                object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];                
                _newNomiDefiniti.AddName(_rigaAttiva, siglaEntitaRif, info["SiglaInformazione"], Struct.tipoVisualizzazione == "O" ? "" : Date.GetSuffissoData(_dataInizio));
                _rigaAttiva++;
            }
        }
        protected virtual void InsertTitoloEntita(object siglaEntita, object desEntita)
        {
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                Range rng = _newNomiDefiniti.Get(siglaEntita, "T", suffissoData);
                rng.Extend(0, _newNomiDefiniti.GetDayOffset(giorno) - 1);

                Excel.Range rngTitolo = _ws.Range[rng.ToString()];
                rngTitolo.Merge();
                rngTitolo.Style = "titleBarStyle";
                rngTitolo.Value = Struct.tipoVisualizzazione == "O" ? desEntita.ToString().ToUpperInvariant() : giorno.ToString("MM/dd/yyyy");
                rngTitolo.RowHeight = 25;
            });
        }
        protected virtual void InsertOre(object siglaEntita)
        {
            if (Struct.tipoVisualizzazione == "V")
            {
                Range rng = _newNomiDefiniti.Get(siglaEntita, "T", Date.GetSuffissoData(_dataInizio));
                rng.StartRow++;
                rng.Extend(0, 25);
                InsertOre(rng);
            }
        }
        private void InsertOre(Range rng, bool hasData0H24 = false)
        {
            Excel.Range rngOre = _ws.Range[rng.ToString()];
            rngOre.Style = "dateBarStyle";
            rngOre.NumberFormat = "0";
            rngOre.Font.Size = 10;
            rngOre.RowHeight = 20;
            
            object[] valoriOre = new object[rng.ColOffset + 1];
            for (int ora = 0; ora < valoriOre.Length; ora++)
            {
                int val = ora + 1;
                if (hasData0H24)
                    val = ora == 0 ? 24 : ora;
                valoriOre[ora] = val;
            }
            rngOre.Value = valoriOre;
        }
        //private void InsertOre(int rigaAttiva, int colonnaInizio, DateTime giorno, int oreGiorno)
        //{
        //    Excel.Range rngOre = _ws.Range[_ws.Cells[rigaAttiva, colonnaInizio], _ws.Cells[rigaAttiva, colonnaInizio + oreGiorno - 1]];
        //    rngOre.Style = "dateBarStyle";
        //    rngOre.NumberFormat = "0";
        //    rngOre.Font.Size = 10;
        //    rngOre.RowHeight = 20;

        //    object[] valoriOre = new object[oreGiorno];
        //    for (int ora = 0; ora < oreGiorno; ora++)
        //    {
        //        int val = ora + 1;
        //        if (giorno == _dataInizio && _struttura.visData0H24)
        //            val = ora == 0 ? 24 : ora;

        //        valoriOre[ora] = val;
        //    }
        //    rngOre.Value = valoriOre;
        //}
        protected virtual void InsertTitoloVerticale(object desEntita)
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;

            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            Range rngTitolo = new Range(_newNomiDefiniti.GetRowByName(siglaEntita, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), _struttura.colBlock - _visParametro - 1, informazioni.Count - 1);

            Excel.Range titoloVert = _ws.Range[rngTitolo.ToString()];
            titoloVert.Style = "titoloVertStyle";
            titoloVert.Merge();
            titoloVert.Orientation = informazioni.Count == 1 ? Excel.XlOrientation.xlHorizontal : Excel.XlOrientation.xlVertical;
            titoloVert.Font.Size = informazioni.Count == 1 ? 6 : 9;

            if (informazioni.Count > 3)
            {
                titoloVert.Value = Struct.tipoVisualizzazione == "O" ? desEntita : _dataInizio;
                titoloVert.NumberFormat = Struct.tipoVisualizzazione == "O" ? "general" : "ddd d";
            }
        }
        protected virtual void FormattaBloccoEntita()
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            Range rng = new Range(_newNomiDefiniti.GetRowByName(siglaEntita, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), _newNomiDefiniti.GetFirstCol() - _visParametro, informazioni.Count - 1, _newNomiDefiniti.GetColOffset(_dataFine) + _visParametro - 1);

            Excel.Range bloccoEntita = _ws.Range[rng.ToString()];
            bloccoEntita.Style = "allDatiStyle";
            bloccoEntita.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            bloccoEntita.Columns[1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            bloccoEntita.Columns[2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            bloccoEntita.Columns[_visParametro].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
            if (_struttura.visParametro)
                bloccoEntita.Columns[3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            int col = _struttura.visData0H24 ? 1 : 0;
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                bloccoEntita.Columns[_visParametro + col].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                col += _newNomiDefiniti.GetDayOffset(suffissoData);
            });
        }
        protected virtual void InsertInformazioniEntita()
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            int col = _newNomiDefiniti.GetFirstCol();
            int colOffset = _newNomiDefiniti.GetColOffset(_dataFine);
            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            int row = _newNomiDefiniti.GetRowByName(siglaEntita, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio));

            Excel.Range rngRow = _ws.Range[Range.GetRange(row, col - _visParametro, informazioni.Count - 1, colOffset + _visParametro - 1)];
            Excel.Range rngInfo = _ws.Range[Range.GetRange(row, col - _visParametro, informazioni.Count - 1, 1)];
            Excel.Range rngData = _ws.Range[Range.GetRange(row, col, informazioni.Count - 1, colOffset - 1)];

            if(Struct.tipoVisualizzazione == "V")
            {
                int oreGiorno = Date.GetOreGiorno(_dataInizio);
                if(oreGiorno < 24)
                    rngData.Columns[rngData.Columns.Count - 1].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                if(oreGiorno < 25)
                    rngData.Columns[rngData.Columns.Count].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
            }

            int i = 1;
            foreach (DataRowView info in informazioni)
            {
                rngInfo.Rows[i].Value = new object[2] { info["DesInformazione"], info["DesInformazioneBreve"] };

                int backColor = (info["BackColor"] is DBNull ? 0 : (int)info["BackColor"]);
                backColor = backColor == 0 || backColor == 2 ? (info["Editabile"].ToString() == "1" ? 15 : 48) : backColor;

                string styleInfo =
                    "FontSize:" + info["FontSize"] + ";" +
                    "ForeColor:" + info["ForeColor"] + ";" +
                    "BackColor:" + backColor + ";" +
                    "Visible:" + info["Visibile"] + ";" +
                    "Borders:[Right:medium];";

                string styleData =
                    "FontSize:" + info["FontSize"] + ";" +
                    "ForeColor:" + info["ForeColor"] + ";" +
                    "Bold:" + info["Grassetto"] + ";" +
                    "NumberFormat:[" + info["Formato"] + "]" +
                    "Align:" + Enum.Parse(typeof(Excel.XlHAlign), info["Align"].ToString());

                if (info["SiglaTipologiaInformazione"].Equals("TITOLO2"))
                {
                    styleInfo += "Merge:true;Borders:[Top:medium]";
                    Style.RangeStyle(rngRow.Rows[i], styleInfo);
                }
                else 
                {
                    if (info["InizioGruppo"].Equals("1"))
                        rngRow.Rows[i].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;

                    Style.RangeStyle(rngInfo.Rows[i], styleInfo);
                    Style.RangeStyle(rngData.Rows[i], styleData);
                    if (info["Data0H24"].Equals("0") && _struttura.visData0H24)
                        rngData.Rows[i].Cells[1].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                }
                i++;
            }
        }
        protected override void InsertPersonalizzazioni() { }
        protected virtual void InsertFormuleValoriDefault()
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            int colOffset = _newNomiDefiniti.GetColOffset(_dataFine);
            foreach (DataRowView info in informazioni)
            {
                object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                
                //tolgo la colonna della DATA0H24 dove non serve
                int offsetAdjust = (_struttura.visData0H24 && info["Data0H24"].Equals("0") ? 1 : 0);
                Range rng = new Range(_newNomiDefiniti.GetRowByName(siglaEntita, info["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), _newNomiDefiniti.GetFirstCol() + offsetAdjust, 0, colOffset - offsetAdjust - 1);

                Excel.Range rngData = _ws.Range[rng.ToString()];
                
                if (info["ValoreDefault"] != DBNull.Value)
                    rngData.Value = info["ValoreDefault"];
                else if (info["FormulaInCella"].Equals("1"))
                {
                    int deltaNeg;
                    int deltaPos;
                    string formula = "=" + PreparaFormula(info, "DATA0", "DATA1", 24, out deltaNeg, out deltaPos);

                    if (info["SiglaTipologiaInformazione"].Equals("OTTIMO"))
                    {
                        rngData.Cells[1].Formula = "=SUM(" + Range.GetRange(rng.StartRow, rng.StartColumn + 1, 0, rng.ColOffset) + ")";
                        deltaNeg = 1;
                    }
                    _ws.Range[Range.GetRange(rng.StartRow, rng.StartColumn + deltaNeg, 0, rng.ColOffset - deltaNeg - deltaPos)].Formula = formula;
                    _ws.Application.ScreenUpdating = false;

                }

                if (info["ValoreData0H24"] != DBNull.Value)
                    rngData.Cells[1].Value = info["ValoreData0H24"];
            }
        }
        protected virtual void InsertParametri()
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            DataView parametriD = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPARAMETROD].DefaultView;
            DataView parametriH = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPARAMETROH].DefaultView;

            foreach (DataRowView info in informazioni)
            {
                object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    Range rngData = _newNomiDefiniti.Get(siglaEntita, info["SiglaInformazione"], suffissoData);
                    rngData.Extend(0, _newNomiDefiniti.GetDayOffset(suffissoData) - 1);

                    Excel.Range rng = _ws.Range[rngData.ToString()];

                    parametriD.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaParametro = '" + info["SiglaTipologiaParametro"] + "' AND CONVERT(DataIV, System.Int32) <= " + giorno.ToString("yyyyMMdd") + " AND CONVERT(DataFV, System.Int32) >= " + giorno.ToString("yyyyMMdd");

                    if (parametriD.Count > 0)
                        rng.Value = parametriD[0]["Valore"];
                    else
                    {
                        parametriH.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaParametro = '" + info["SiglaTipologiaParametro"] + "' AND CONVERT(DataIV, System.Int32) <= " + giorno.ToString("yyyyMMdd") + " AND CONVERT(DataFV, System.Int32) >= " + giorno.ToString("yyyyMMdd");

                        parametriH.Sort = "Ora";

                        object[] values = parametriH.ToTable(false, "Valore").AsEnumerable().Select(r => r["Valore"]).ToArray();
                        
                        if(values.Length > 0)
                            rng.Value = values;
                    }
                });
            }
        }
        protected virtual void FormattazioneCondizionale()
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            DataView formattazione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONEFORMATTAZIONE].DefaultView;
            int colOffset = _newNomiDefiniti.GetColOffset(_dataFine);
            foreach (DataRowView info in informazioni)
            {
                object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                
                int offsetAdjust = (_struttura.visData0H24 && info["Data0H24"].Equals("0") ? 1 : 0);
                Range rng = new Range(_newNomiDefiniti.GetRowByName(siglaEntita, info["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), _newNomiDefiniti.GetFirstCol() + offsetAdjust, 0, colOffset - offsetAdjust - 1);

                Excel.Range rngData = _ws.Range[rng.ToString()];

                formattazione.RowFilter = (info["SiglaEntitaRif"] is DBNull ? "SiglaEntita" : "SiglaEntitaRif") + " = '" + siglaEntita + "' AND SiglaInformazione = '" + info["SiglaInformazione"] + "'";
                foreach (DataRowView format in formattazione)
                {
                    string[] valore = format["Valore"].ToString().Replace("\"", "").Split('|');
                    if (format["NomeCella"] != DBNull.Value)
                    {
                        int refRow = _newNomiDefiniti.GetRowByName(siglaEntita, format["NomeCella"], Struct.tipoVisualizzazione == "O" ? "" : Date.GetSuffissoData(_dataInizio));
                        string address = Range.GetRange(refRow, rng.StartColumn);
                        string formula = "";
                        switch ((int)format["Operatore"])
                        {
                            case 1:
                                formula = "=E(" + address + ">=" + valore[0] + ";" + address + "<=" + valore[1] + ")";
                                break;
                            case 3:
                                formula = "=" + address + "=" + valore[0];
                                break;
                            case 5:
                                formula = "=" + address + ">" + valore[0];
                                break;
                            case 6:
                                formula = "=" + address + "<" + valore[0];
                                break;
                        }
                        Excel.FormatCondition cond = rngData.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: formula);

                        cond.Font.Color = format["ForeColor"];
                        cond.Font.Bold = format["Grassetto"].Equals("1");
                        if ((int)format["BackColor"] == 0)
                            cond.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                        else
                            cond.Interior.Color = format["BackColor"];
                        cond.Interior.Pattern = format["Pattern"];
                    }
                    else
                    {
                        string formula1;
                        string formula2 = "";
                        if ((int)format["Operatore"] == 1)
                        {
                            formula1 = valore[0];
                            formula2 = valore[1];
                        }
                        else
                        {
                            formula1 = valore[0];
                        }

                        Excel.FormatCondition cond = rngData.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, format["Operatore"], formula1, formula2);

                        cond.Font.Color = format["ForeColor"];
                        cond.Font.Bold = format["Grassetto"].Equals("1");
                        if ((int)format["BackColor"] == 0)
                            cond.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                        else
                            cond.Interior.Color = format["BackColor"];

                        cond.Interior.Pattern = format["Pattern"];
                    }
                }
            }
        }
        protected virtual void InsertGrafici()
        {
            DataView grafici = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAGRAFICO].DefaultView;
            DataView graficiInfo = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAGRAFICOINFORMAZIONE].DefaultView;

            int i = 1;
            int col = _newNomiDefiniti.GetFirstCol() + (_struttura.visData0H24 ? 1 : 0);
            int colOffset = _newNomiDefiniti.GetColOffset(_dataFine) - (_struttura.visData0H24 ? 1 : 0);
            foreach (DataRowView grafico in grafici)
            {
                string name = NewDefinedNames.GetName(grafico["SiglaEntita"], "GRAFICO" + i++, Struct.tipoVisualizzazione == "O" ? "" : Date.GetSuffissoData(_dataInizio));

                Range rngGrafico = new Range(_newNomiDefiniti.GetRowByName(name), col, 0, colOffset - 1);
                //int row = _newNomiDefiniti.GetRowByName(name);
                Excel.Range xlRngGrafico = _ws.Range[rngGrafico.ToString()];
                xlRngGrafico.Merge();
                xlRngGrafico.Style = "chartsBarStyle";
                xlRngGrafico.RowHeight = 200;
                Excel.Chart chart = _ws.ChartObjects().Add(xlRngGrafico.Left, xlRngGrafico.Top + 1, xlRngGrafico.Width, xlRngGrafico.Height - 2).Chart;

                chart.Parent.Name = name;

                chart.Axes(Excel.XlAxisType.xlCategory).TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
                chart.Axes(Excel.XlAxisType.xlValue).HasMajorGridlines = false;
                chart.Axes(Excel.XlAxisType.xlValue).HasMinorGridlines = false;
                chart.Axes(Excel.XlAxisType.xlValue).MinorTickMark = Excel.XlTickMark.xlTickMarkOutside;
                chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Name = "Verdana";
                chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Size = 11;
                chart.Axes(Excel.XlAxisType.xlValue).TickLabels.NumberFormat = "general";

                chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;
                chart.HasDataTable = false;
                chart.DisplayBlanksAs = Excel.XlDisplayBlanksAs.xlNotPlotted;
                chart.ChartGroups(1).GapWidth = 0;
                chart.ChartGroups(1).Overlap = 100;
                chart.ChartArea.Border.ColorIndex = 1;
                chart.ChartArea.Border.Weight = 3;
                chart.ChartArea.Border.LineStyle = 0;

                chart.PlotArea.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                foreach (DataRowView info in graficiInfo)
                {
                    Range rngDati = new Range(_newNomiDefiniti.GetRowByName(grafico["SiglaEntita"], info["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), col, 0, colOffset - 1);
                    Excel.Range datiGrafico = _ws.Range[rngDati.ToString()];
                    var serie = chart.SeriesCollection().Add(datiGrafico);
                    serie.Name = info["DesInformazione"].ToString();
                    serie.ChartType = (Excel.XlChartType)info["ChartType"];
                    serie.Interior.ColorIndex = info["InteriorColor"];
                    serie.Border.ColorIndex = info["BorderColor"];
                    serie.Border.Weight = info["BorderWeight"];
                    serie.Border.LineStyle = info["BorderLineStyle"];
                }
            }
            _ws.Application.ScreenUpdating = false;
        }
        public override void AggiornaGrafici()
        {
            //_ws.Application.CalculateFull();
            //Excel.ChartObjects charts = _ws.ChartObjects();
            //foreach (Excel.ChartObject chart in charts)
            //{
            //    int col;
            //    if(chart.Name.Contains("DATA")) 
            //    {
            //        col = _newNomiDefiniti.GetColFromDate(chart.Name.Split(Simboli.UNION[0]).Last());
            //    } 
            //    else 
            //    {
            //        col = _newNomiDefiniti.GetColFromDate();
            //    }
            //    int row = _newNomiDefiniti.GetRowByName(chart.Name);
            //    Excel.Range rng = _ws.Range[Range.GetRange(row, col)];
            //    AggiornaGrafici(chart.Chart, rng.MergeArea);
            //    chart.Chart.Refresh();
            //}
        }
        private void AggiornaGrafici(Excel.Chart chart, Excel.Range rigaGrafico)
        {
            //resize dell'area del grafico per adattarla alle ore
            string max = chart.Axes(Excel.XlAxisType.xlValue).MaximumScale.ToString();
            string min = chart.Axes(Excel.XlAxisType.xlValue).MinimumScale.ToString();

            max = max.Length > min.Length ? max : min;

            Graphics grfx = Graphics.FromImage(new Bitmap(1, 1));
            grfx.PageUnit = GraphicsUnit.Point;
            SizeF sizeMax = grfx.MeasureString(max, new Font("Verdana", 11));

            chart.ChartArea.Left = rigaGrafico.Left - sizeMax.Width - 7;
            chart.ChartArea.Width = rigaGrafico.Width + sizeMax.Width + 4;
            chart.PlotArea.InsideLeft = 0;
            chart.PlotArea.Width = chart.ChartArea.Width + 3;
        }

        #endregion

        public override void CaricaInformazioni(bool all)
        {
            try
            {
                DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
                DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPROPRIETA].DefaultView;
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

                DataView datiApplicazione = DataBase.DB.Select(DataBase.SP.APPLICAZIONE_INFORMAZIONE_H, "@SiglaCategoria=" + _siglaCategoria + ";@SiglaEntita=ALL;@DateFrom=" + _dataInizio.ToString("yyyyMMdd") + ";@DateTo=" + dataFineMax.ToString("yyyyMMdd") + ";@Tipo=1;@All=" + (all ? "1" : "0")).DefaultView;

                DataView insertManuali = new DataView();
                if (all)
                    insertManuali = DataBase.DB.Select(DataBase.SP.APPLICAZIONE_INFORMAZIONE_COMMENTO, "@SiglaCategoria=" + _siglaCategoria + ";@SiglaEntita=ALL;@DateFrom=" + _dataInizio.ToString("yyyyMMdd") + ";@DateTo=" + dataFineMax.ToString("yyyyMMdd") + ";@All=1").DefaultView;

                if (Struct.tipoVisualizzazione == "O")
                {
                    foreach (DataRowView entita in categoriaEntita)
                    {

                        datiApplicazione.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND CONVERT(Data, System.Int32) <= " + dateFineUP[entita["SiglaEntita"]].ToString("yyyyMMdd");

                        //_dataFine = dateFineUP[entita["SiglaEntita"]];
                        CaricaInformazioniEntita(datiApplicazione);
                        if (all)
                        {
                            insertManuali.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND CONVERT(SUBSTRING(Data, 1, 8), System.Int32) <= " + dateFineUP[entita["SiglaEntita"]].ToString("yyyyMMdd");
                            CaricaCommentiEntita(insertManuali);
                        }
                    }
                }
                else
                {
                    CaricaInformazioniEntita(datiApplicazione);
                    if (all)
                    {
                        CaricaCommentiEntita(insertManuali);
                    }
                }
            }
            catch (Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "CaricaInformazioni [all = " + all + "]: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        private void CaricaInformazioniEntita(DataView datiApplicazione)
        {
            foreach (DataRowView dato in datiApplicazione)
            {
                DateTime giorno = DateTime.ParseExact(dato["Data"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                //sono nel caso DATA0H24
                if (giorno < DataBase.DataAttiva)
                {
                    Range rng = _newNomiDefiniti.Get(dato["SiglaEntita"], dato["SiglaInformazione"], Date.GetSuffissoData(DataBase.DataAttiva.AddDays(-1)), Date.GetSuffissoOra(24));
                    _ws.Range[rng.ToString()].Value = dato["H24"];
                }
                else
                {
                    //int col = Struct.tipoVisualizzazione == "O" ? _newNomiDefiniti.GetColFromDate(giorno) : _newNomiDefiniti.GetFirstCol();
                    //int dayOffset = Struct.tipoVisualizzazione == "O" ? _newNomiDefiniti.GetDayOffset(giorno) : _newNomiDefiniti.GetColOffset();
                    //int row = _newNomiDefiniti.GetRowByName(dato["SiglaEntita"], dato["SiglaInformazione"], Struct.tipoVisualizzazione == "O" ? "" : Date.GetSuffissoData(giorno));

                    Range rng = _newNomiDefiniti.Get(dato["SiglaEntita"], dato["SiglaInformazione"], Date.GetSuffissoData(giorno));
                    rng.Extend(0, _newNomiDefiniti.GetDayOffset(giorno) - 1);

                    List<object> o = new List<object>(dato.Row.ItemArray);
                    //elimino i campi inutili
                    o.RemoveRange(o.Count - 3, 3);
                    _ws.Range[rng.ToString()].Value = o.ToArray();
                }
            }
        }
        private void CaricaCommentiEntita(DataView insertManuali)
        {
            foreach (DataRowView commento in insertManuali)
            {
                Range rngComm = _newNomiDefiniti.Get(commento["SiglaEntita"], commento["SiglaInformazione"], Date.GetSuffissoData(commento["Data"].ToString()), Date.GetSuffissoOra(commento["Data"].ToString()));

                Excel.Range rng = _ws.Range[rngComm.ToString()];
                rng.ClearComments();
                rng.AddComment("Valore inserito manualmente");
            }
        }

        //TODO
        public override void CalcolaFormule(string siglaEntita = null, DateTime? giorno = null, int ordineElaborazione = 0, bool escludiOrdine = false)
        {
            DataView dvCE = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            DataView dvEP = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPROPRIETA].DefaultView;
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;

            dvCE.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL )" + (siglaEntita == null ? "" : " AND SiglaEntita = '" + siglaEntita + "'");

            //_dataInizio = DB.DataAttiva;
            //DateTime giorno = dataAttiva ?? DB.DataAttiva;

            bool all = giorno == null;

            foreach (DataRowView entita in dvCE)
            {
                siglaEntita = entita["SiglaEntita"].ToString();

                informazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND OrdineElaborazione <> 0 AND FormulaInCella = 0";
                if (ordineElaborazione != 0)
                {
                    informazioni.RowFilter += " AND OrdineElaborazione" + (escludiOrdine ? " <> " : " = ") + ordineElaborazione;
                }
                informazioni.Sort = "OrdineElaborazione";

                if (informazioni.Count > 0)
                {
                    DateTime dataFine;

                    dvEP.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                    if (dvEP.Count > 0)
                        dataFine = DataBase.DB.DataAttiva.AddDays(double.Parse("" + dvEP[0]["Valore"]));
                    else
                        dataFine = DataBase.DB.DataAttiva.AddDays(Struct.intervalloGiorni);

                    string suffissoData = all ? "DATA1" : Date.GetSuffissoData(DataBase.DB.DataAttiva, giorno.Value);
                    string suffissoDataPrec = all ? "DATA0" : Date.GetSuffissoData(DataBase.DB.DataAttiva, giorno.Value.AddDays(-1));
                    string suffissoUltimoGiorno = Date.GetSuffissoData(DataBase.DB.DataAttiva, dataFine);

                    foreach (DataRowView info in informazioni)
                    {
                        Tuple<int, int>[] riga;
                        if (all)
                            riga = _nomiDefiniti[info["Data0H24"].Equals("0"), entita["SiglaEntita"], info["SiglaInformazione"]];
                        else
                            riga = _nomiDefiniti[entita["SiglaEntita"], info["SiglaInformazione"], suffissoData];


                        int deltaNeg;
                        int deltaPos;
                        int oreDataPrec = all ? 24 : Date.GetOreGiorno(giorno.Value.AddDays(-1));

                        string formula = "=" + PreparaFormula(info, suffissoDataPrec, suffissoData, oreDataPrec, out deltaNeg, out deltaPos);

                        if (suffissoData != "DATA1")
                            deltaNeg = 0;
                        if (suffissoData != suffissoUltimoGiorno)
                            deltaPos = 0;

                        Excel.Range rng = _ws.Range[_ws.Cells[riga[0].Item1, riga[0].Item2 - deltaNeg], _ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2 - deltaPos]];

                        rng.Formula = formula;
                    }
                }
                informazioni.Sort = "";
            }

        }

        protected string PreparaFormula(DataRowView info, string suffissoDataPrec, string suffissoData, int oreDataPrec, out int deltaNeg, out int deltaPos)
        {
            if (info["Formula"] != DBNull.Value || info["Funzione"] != DBNull.Value)
            {
                string formula = info["Formula"] is DBNull ? info["Funzione"].ToString() : info["Formula"].ToString();

                string[] parametri = info["FormulaParametro"].ToString().Split(',');

                int tmpdeltaNeg = 0;
                int tmpdeltaPos = 0;

                foreach (string par in parametri)
                {
                    if (Regex.IsMatch(par, @"\[[-+]?\d+\]"))
                    {
                        int deltaOre = int.Parse(par.Split('[')[1].Replace("]", ""));
                        if (deltaOre > 0)
                            tmpdeltaPos = Math.Max(tmpdeltaPos, deltaOre);
                        else
                            tmpdeltaNeg = Math.Min(tmpdeltaNeg, deltaOre);
                    }
                }

                deltaNeg = Math.Abs(tmpdeltaNeg);
                deltaPos = tmpdeltaPos;

                formula = Regex.Replace(formula, @"%P\d+(E\d+)?%", delegate(Match m)
                {
                    string[] parametroEntita = m.Value.Split('E');
                    int n = int.Parse(Regex.Match(parametroEntita[0], @"\d+").Value);

                    object siglaEntita = "";
                    string siglaInformazione = "";
                    string suffData = "";
                    string suffOra = "";
                    if (parametroEntita.Length > 1)
                    {
                        int eRif = int.Parse(Regex.Match(parametroEntita[1], @"\d+").Value);
                        DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
                        categoriaEntita.RowFilter = "Gerarchia = '" + info["SiglaEntita"] + "' AND Riferimento = " + eRif;
                        siglaEntita = categoriaEntita[0]["SiglaEntita"];
                    }
                    else
                        siglaEntita = info["SiglaEntita"];
                    
                    siglaInformazione = parametri[n - 1];

                    if (Regex.IsMatch(siglaInformazione, @"\[[-+]?\d+\]"))
                    {
                        int deltaOre = int.Parse(siglaInformazione.Split('[')[1].Replace("]", ""));

                        if (suffissoData == "DATA1")
                        {//traslo in avanti la formula di |deltaNeg| - |deltaOre|
                            int ora = Math.Abs(tmpdeltaNeg) + deltaOre + (info["Data0H24"].Equals("1") ? 0 : 1);
                            suffData = ora == 0 ? "DATA0" : "DATA1";
                            suffOra = ora == 0 ? "H24" : "H" + ora;
                        }
                        else
                        {
                            int ora = (deltaOre < 0 ? oreDataPrec + deltaOre + 1 : deltaOre + 1);
                            suffData = deltaOre < 0 ? suffissoDataPrec : suffissoData;
                            suffOra = "H" + ora;
                        }
                        siglaInformazione = Regex.Replace(siglaInformazione, @"\[[-+]?\d+\]", "");
                    }
                    else
                    {
                        if (suffissoData == "DATA1")
                        {
                            int ora = tmpdeltaNeg == 0 ? 1 : Math.Abs(tmpdeltaNeg) + (info["Data0H24"].Equals("1") ? 0 : 1);
                            suffData = suffissoData;
                            suffOra = "H" + ora;
                        }
                        else
                        {
                            suffData = suffissoData;
                            suffOra = "H1";
                        }
                    }
                    Range rng = _newNomiDefiniti.Get(siglaEntita, siglaInformazione, suffData, suffOra);

                    return rng.ToString();
                }, RegexOptions.IgnoreCase);
                return formula;
            }
            deltaNeg = 0;
            deltaPos = 0;

            return "";
        }

        public override void UpdateData(bool all = true)
        {
            if (all)
            {
                CancellaDati();
                AggiornaDateTitoli();
                CaricaParametri();
                //CancellaDati();
                //AggiornaDateTitoli();
                //CaricaParametri();
            }
            CaricaInformazioni(all);
            AggiornaGrafici();
        }
        #region UpdateData

        private void CancellaDati()
        {
            CancellaDati(DataBase.DataAttiva, true);
        }
        private void CancellaDati(DateTime giorno, bool all = false)
        {
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPROPRIETA].DefaultView;
            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "'"; // AND (Gerarchia = '' OR Gerarchia IS NULL )";

            string suffissoData = Date.GetSuffissoData(giorno);
            int colOffset = _newNomiDefiniti.GetColOffset();
            if (!all)
                colOffset = _newNomiDefiniti.GetColOffset(suffissoData);

            foreach (DataRowView entita in categoriaEntita)
            {
                DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND FormulaInCella = '0' AND ValoreDefault IS NULL";

                foreach (DataRowView info in informazioni)
                {
                    int col = all ? _newNomiDefiniti.GetFirstCol() : _newNomiDefiniti.GetColFromDate(suffissoData);
                    object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                    if (Struct.tipoVisualizzazione == "O")
                    {
                        int row = _newNomiDefiniti.GetRowByName(siglaEntita, info["SiglaInformazione"]);
                        Excel.Range rngData = _ws.Range[Range.GetRange(row, col, 0, colOffset - 1)];
                        rngData.Value = "";
                        rngData.ClearComments();
                        Style.RangeStyle(rngData, "BackColor:" + info["BackColor"] + ";ForeColor:" + info["ForeColor"]);
                    }
                    else
                    {
                        DateTime dataInizio = giorno;
                        DateTime dataFine = giorno;
                        if(all)
                        {
                            dataInizio = DataBase.DataAttiva;
                            dataFine = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni);
                        }

                        CicloGiorni(dataInizio, dataFine, (oreGiorno, suffData, g) =>
                        {
                            int row = _newNomiDefiniti.GetRowByName(siglaEntita, info["SiglaInformazione"], suffData);

                            Excel.Range rngData = _ws.Range[Range.GetRange(row, col, 0, colOffset - 1)];
                            rngData.Value = "";
                            rngData.ClearComments();
                            Style.RangeStyle(rngData, "BackColor:" + info["BackColor"] + ";ForeColor:" + info["ForeColor"]);
                        });
                    }
                }
            }
        }
        public override void AggiornaDateTitoli()
        {
            if (Struct.tipoVisualizzazione == "O")
            {
                int row = _struttura.rigaBlock - 2;
                for (int i = 0; i < _newNomiDefiniti.DaySuffx.Length; i++)
                {
                    if (_newNomiDefiniti.DaySuffx[i] != "DATA0")
                    {
                        int col = _newNomiDefiniti.GetColFromDate(_newNomiDefiniti.DaySuffx[i]);
                        _ws.Range[Range.GetRange(row, col)].Value = Date.GetDataFromSuffisso(_newNomiDefiniti.DaySuffx[i]);
                    }
                }
            }
            else
            {
                for (int i = 0; i < Struct.intervalloGiorni; i++)
                {
                    DateTime giorno = DataBase.DataAttiva.AddDays(i);
                    string suffissoData = Date.GetSuffissoData(giorno);
                    
                    int row = _newNomiDefiniti.GetRowByName(suffissoData, "T");
                    int col = _newNomiDefiniti.GetFirstCol();
                    _ws.Range[Range.GetRange(row, col)].Value = giorno;
                    
                    row = _newNomiDefiniti.GetRowByName(suffissoData, "TITOLO_VERTICALE");
                    col -= (_visParametro + 1);
                    _ws.Range[Range.GetRange(row, col)].Value = giorno;

                }
            }
        }
        protected void CaricaParametri()
        {
            DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPROPRIETA].DefaultView;
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;

            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";
            _dataInizio = DataBase.DB.DataAttiva;

            foreach (DataRowView entita in categoriaEntita)
            {
                entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                if (entitaProprieta.Count > 0)
                    _dataFine = _dataInizio.AddDays(double.Parse("" + entitaProprieta[0]["Valore"]));
                else
                    _dataFine = _dataInizio.AddDays(Struct.intervalloGiorni);

                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaTipologiaParametro IS NOT NULL";
                InsertParametri();
            }
        }

        #endregion

        #endregion

        public void Dispose()
        {
            if (!_disposed)
            {
                GC.SuppressFinalize(this);
                _disposed = true;
            }
        }
    }
}
