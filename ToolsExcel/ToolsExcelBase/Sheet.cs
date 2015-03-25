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
        protected abstract void InsertPersonalizzazioni(object siglaEntita);

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
        public static string R1C1toA1(int riga, int colonna)
        {
            string output = "";
            while (colonna > 0)
            {
                int lettera = (colonna - 1) % 26;
                output = Convert.ToChar(lettera + 65) + output;
                colonna = (colonna - lettera) / 26;
            }
            output += riga;
            return output;
        }
        public static string R1C1toA1(Tuple<int,int> cella)
        {
            return R1C1toA1(cella.Item1, cella.Item2);
        }
        public static Tuple<int, int> A1toR1C1(string address)
        {
            address = address.Replace("$", "");
            string alpha = Regex.Match(address, @"\D+").Value;
            int riga = int.Parse(Regex.Match(address, @"\d+").Value);

            int colonna = 0;
            int incremento = (alpha.Length == 1 ? 1 : 26 * (alpha.Length - 1));
            for (int i = 0; i < alpha.Length; i++)
            {
                colonna += (char.ConvertToUtf32(alpha, i) - 64) * incremento;
                incremento = incremento - 26 == 0 ? 1 : incremento - 26;
            }

            return Tuple.Create<int, int>(riga, colonna);
        }

        public static string GetRange(int startRow, int startCol, int rowOffset = 0, int colOffset = 0)
        {
            if (rowOffset == 0 && colOffset == 0)
                return R1C1toA1(startRow, startCol);

            return R1C1toA1(startRow, startCol) + ":" + R1C1toA1(startRow + rowOffset, startCol + colOffset);
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
                                Handler.StoreEdit(ws, ws.Range[nomiDefiniti.GetRange(rngInfo)]);
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
                            Handler.StoreEdit(ws, ws.Range[nomiDefiniti.GetRange(rngInfo)]);
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

        protected Cell _cell;
        

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

            _cell = new Cell();
            _struttura = new Struct();

            _cell.Width.empty = double.Parse(paramApplicazione[0]["ColVuotaWidth"].ToString());
            _cell.Width.dato = double.Parse(paramApplicazione[0]["ColDatoWidth"].ToString());
            _cell.Width.entita = double.Parse(paramApplicazione[0]["ColEntitaWidth"].ToString());
            _cell.Width.informazione = double.Parse(paramApplicazione[0]["ColInformazioneWidth"].ToString());
            _cell.Width.unitaMisura = double.Parse(paramApplicazione[0]["ColUMWidth"].ToString());
            _cell.Width.parametro = double.Parse(paramApplicazione[0]["ColParametroWidth"].ToString());
            _cell.Width.jolly1 = double.Parse(paramApplicazione[0]["ColJolly1Width"].ToString());
            _cell.Height.normal = double.Parse(paramApplicazione[0]["RowHeight"].ToString());
            _cell.Height.empty = double.Parse(paramApplicazione[0]["RowVuotaHeight"].ToString());

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
                        intervalloGiorniMax = Math.Max(intervalloGiorniMax, int.Parse("" + entitaProprieta[0]["Valore"]));
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

            //CaricaInformazioni(all: true);
            //CalcolaFormule();
            //Utilities.AggiornaFormule(_ws);
            //InsertGrafici();

            _newNomiDefiniti.DumpToDataSet();
        }

        protected void Clear()
        {
            _ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            _ws.UsedRange.EntireColumn.Delete();
            _ws.UsedRange.FormatConditions.Delete();
            _ws.UsedRange.EntireRow.Hidden = false;
            _ws.UsedRange.Font.Size = 10;
            _ws.UsedRange.NumberFormat = "General";
            _ws.UsedRange.Font.Name = "Verdana";
            _ws.UsedRange.RowHeight = _cell.Height.normal;

            _ws.Rows["1:" + (_struttura.rigaBlock - 1)].RowHeight = _cell.Height.empty;
            _ws.Rows[_struttura.rigaGoto].RowHeight = _cell.Height.normal;

            _ws.Columns[1].ColumnWidth = _cell.Width.empty;
            _ws.Columns[2].ColumnWidth = _cell.Width.entita;

            ((Excel._Worksheet)_ws).Activate();
            _ws.Application.ActiveWindow.FreezePanes = false;
            _ws.Cells[_struttura.rigaBlock, _struttura.colBlock].Select();
            _ws.Application.ActiveWindow.ScrollColumn = 1;
            _ws.Application.ActiveWindow.ScrollRow = 1;
            _ws.Application.ActiveWindow.FreezePanes = true;

            int colInfo = _struttura.colBlock - _visParametro;
            _ws.Columns[colInfo].ColumnWidth = _cell.Width.informazione;
            _ws.Columns[colInfo + 1].ColumnWidth = _cell.Width.unitaMisura;
            if (_struttura.visParametro)
                _ws.Columns[colInfo + 2].ColumnWidth = _cell.Width.parametro;
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

                object nome = Struct.tipoVisualizzazione == "O" ? categoriaEntita[i]["SiglaEntita"] : DataBase.DataAttiva.AddDays(i);
                _newNomiDefiniti.AddGOTO(nome, _struttura.rigaGoto + r, _struttura.colBlock + c);

                Excel.Range rng;
                if (_cell.Width.dato < 8)
                {
                    j = c == 0 ? 0 : j + 1;
                    c += j;
                    rng = _ws.Range[_ws.Cells[_struttura.rigaGoto + r, _struttura.colBlock + c], _ws.Cells[_struttura.rigaGoto + r, _struttura.colBlock + c + 1]];
                    rng.Merge();
                }
                else
                {
                    rng = _ws.Cells[_struttura.rigaGoto + r, _struttura.colBlock + c];
                }
                rng.Value = Struct.tipoVisualizzazione == "O" ? categoriaEntita[i]["DesEntitaBreve"] : nome;
                rng.Style = Struct.tipoVisualizzazione == "O" ? "navBarStyleHorizontal" : "navBarStyleVertical";
            }

            //inserisco la data e le ore
            if (Struct.tipoVisualizzazione == "O")
            {
                int colonnaInizio = _struttura.colBlock;
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    int escludiH24 = (giorno == _dataInizio && _struttura.visData0H24 ? 1 : 0);
                    Excel.Range rngData = _ws.Range[_ws.Cells[_struttura.rigaBlock - 2, colonnaInizio + escludiH24], _ws.Cells[_struttura.rigaBlock - 2, colonnaInizio + oreGiorno - 1]];
                    rngData.Merge();
                    rngData.Style = "dateBarStyle";
                    rngData.Value = giorno.ToString("MM/dd/yyyy");
                    rngData.RowHeight = 20;

                    InsertOre(_struttura.rigaBlock - 1, colonnaInizio, giorno, oreGiorno);
                    colonnaInizio += oreGiorno;
                });
            }
        }
        
        protected void InitBloccoEntita(DataRowView entita)
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            DataView grafici = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAGRAFICO].DefaultView;
            
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
            informazioni.Sort = "Ordine";

            grafici.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

            _colonnaInizio = _struttura.colBlock;
            _intervalloOre = Date.GetOreIntervallo(_dataInizio, _dataFine) + (_struttura.visData0H24 ? 1 : 0) + (_struttura.visParametro ? 1 : 0);

            CreaNomiCelle2(entita["SiglaEntita"]);
            InsertTitoloEntita2(entita["SiglaEntita"], entita["DesEntita"]);
            InsertOre2(entita["SiglaEntita"]);
            InsertTitoloVerticale2(entita["SiglaEntita"], entita["DesEntitaBreve"]);
            FormattaBloccoEntita(entita["SiglaEntita"]);
            InsertInformazioniEntita2();

            //InsertTitoloEntita(entita);
            //InsertRangeGrafici(entita["SiglaEntita"]);
            //InsertOre(entita["SiglaEntita"]);
            //InsertTitoloVerticale(entita["SiglaEntita"], entita["DesEntitaBreve"], informazioni.Count);
            //FormattaAllDati(entita["SiglaEntita"]);
            //InsertInformazioniEntita(entita["SiglaEntita"]);
            //CreaNomiCelle(entita["SiglaEntita"]);
            //InsertPersonalizzazioni(entita["SiglaEntita"]);
            //InsertValoriCelle(entita["SiglaEntita"]);
            //InsertParametri(entita["SiglaEntita"]);
            //CreaFormattazioneCondizionale(entita["SiglaEntita"]);

            //due righe vuote tra un'entità e la successiva
            _rigaAttiva += 2;
        }
        #region Blocco entità

        protected virtual void CreaNomiCelle2(object siglaEntita)
        {
            //inserisco titoli
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                _newNomiDefiniti.AddName(_rigaAttiva, siglaEntita, "T", suffissoData);
                //sistema l'indirizzamento dei GOTO
                int col = _newNomiDefiniti.GetColFromDate(suffissoData);
                object name = Struct.tipoVisualizzazione == "O" ? siglaEntita : giorno;
                _newNomiDefiniti.ChangeGOTOAddressTo(name, Sheet.R1C1toA1(_rigaAttiva, col));
            });

            //aggiungo la riga delle ore
            _rigaAttiva += Struct.tipoVisualizzazione == "V" ? 2 : 1;

            //aggiungo i grafici
            DataView grafici = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAGRAFICO].DefaultView;

            int i = 1;
            foreach (DataRowView grafico in grafici)
            {
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    _newNomiDefiniti.AddName(_rigaAttiva, grafico["SiglaEntita"], "GRAFICO" + i, suffissoData);
                });
                i++;
                _rigaAttiva++;
            }

            //aggiungo informazioni
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            foreach (DataRowView info in informazioni)
            {
                object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                _newNomiDefiniti.AddName(_rigaAttiva, siglaEntitaRif, info["SiglaInformazione"], Date.GetSuffissoData(_dataInizio));
                _rigaAttiva++;
            }
        }
        protected virtual void InsertTitoloEntita2(object siglaEntita, object desEntita)
        {
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                int col = _newNomiDefiniti.GetColFromDate(suffissoData);
                int row = _newNomiDefiniti.GetRowByName(siglaEntita, "T", suffissoData);
                //TODO usare oreGiorno (senza modifica) una volta che tolgo la condizione che aggiunge uno se sono nel primo giorno e ho data0h24
                oreGiorno = Struct.tipoVisualizzazione == "O" ? Date.GetOreGiorno(giorno) : 25;
                Excel.Range rngTitolo = _ws.Range[GetRange(row, col, 0, oreGiorno - 1)];
                rngTitolo.Merge();
                rngTitolo.Style = "titleBarStyle";
                rngTitolo.Value = Struct.tipoVisualizzazione == "O" ? desEntita.ToString().ToUpperInvariant() : giorno.ToString("MM/dd/yyyy");
                rngTitolo.RowHeight = 25;
            });
        }
        protected virtual void InsertOre2(object siglaEntita)
        {
            if (Struct.tipoVisualizzazione == "V")
            {
                int col = _newNomiDefiniti.GetColFromDate();
                int row = _newNomiDefiniti.GetRowByName(siglaEntita, "T", Date.GetSuffissoData(_dataInizio)) + 1;

                InsertOre(row, col, _dataInizio, 25);
            }
        }
        protected virtual void InsertTitoloVerticale2(object siglaEntita, object desEntita)
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            
            object siglaEntitaRif = informazioni[0]["SiglaEntitaRif"] is DBNull ? siglaEntita : informazioni[0]["SiglaEntitaRif"];
            int row = _newNomiDefiniti.GetRowByName(siglaEntitaRif, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio));

            Excel.Range titoloVert = _ws.Range[GetRange(row, _struttura.colBlock - _visParametro - 1, informazioni.Count - 1)];
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
        protected virtual void FormattaBloccoEntita(object siglaEntita)
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            int col = _newNomiDefiniti.GetFirstCol();
            int colOffset = _newNomiDefiniti.GetColOffset();
            object siglaEntitaRif = informazioni[0]["SiglaEntitaRif"] is DBNull ? siglaEntita : informazioni[0]["SiglaEntitaRif"];
            int row = _newNomiDefiniti.GetRowByName(siglaEntitaRif, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio));

            Excel.Range bloccoEntita = _ws.Range[GetRange(row, col - _visParametro, informazioni.Count - 1, colOffset + _visParametro - 1)];
            bloccoEntita.Style = "allDatiStyle";
            bloccoEntita.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            bloccoEntita.Columns[1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            bloccoEntita.Columns[2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            bloccoEntita.Columns[_visParametro].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
            if (_struttura.visParametro)
                bloccoEntita.Columns[3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            col = _struttura.visData0H24 ? 1 : 0;
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                bloccoEntita.Columns[_visParametro + col].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                col += _newNomiDefiniti.GetColOffset(suffissoData);
            });
        }
        protected virtual void InsertInformazioniEntita2()
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            int col = _struttura.colBlock - _visParametro;
            int colOffset = _newNomiDefiniti.GetFirstCol() - col + _newNomiDefiniti.GetColOffset();
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                suffissoData = Struct.tipoVisualizzazione == "O" ? Date.GetSuffissoData(DataBase.DataAttiva) : suffissoData;
                object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
                int row = _newNomiDefiniti.GetRowByName(siglaEntita, informazioni[0]["SiglaInformazione"], suffissoData);

                Excel.Range rngInfo = _ws.Range[GetRange(row, col, informazioni.Count - 1, 1)];

                int i = 1;
                foreach (DataRowView info in informazioni)
                {
                    rngInfo.Rows[i].Value = new object[2] { info["DesInformazione"], info["DesInformazioneBreve"] };

                    int backColor = (info["BackColor"] is DBNull ? 0 : (int)info["BackColor"]);
                    backColor = backColor == 0 || backColor == 2 ? (info["Editabile"].ToString() == "1" ? 15 : 48) : backColor;

                    string style = "FontSize:" + info["FontSize"] + ";FontName:Verdana;BackColor:" + backColor + ";"
                    + "ForeColor:" + info["ForeColor"] + ";Visible:" + info["Visibile"] + ";";

                    if (info["SiglaTipologiaInformazione"].Equals("TITOLO2"))
                    {
                        style += "Bold:" + info["Grassetto"] + ";Merge:true;Borders:[Top:medium]";
                        Style.RangeStyle(_ws.Range[GetRange(row + i - 1, col, 0, colOffset - 1)], style);
                    }
                    else if (info["InizioGruppo"].Equals("1"))
                    {
                        _ws.Range[GetRange(row + i - 1, col, 0, colOffset - 1)].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
                    }
                    Style.RangeStyle(rngInfo.Rows[i++], style);
                }
            });

        }

        protected virtual void InsertTitoloEntita(DataRowView entita)
        {
            int colonnaInizio = _colonnaInizio;            
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                bool isVisibleData0H24 = giorno == _dataInizio && _struttura.visData0H24;

                if (isVisibleData0H24)
                {
                    colonnaInizio++;
                    oreGiorno--;
                }
                Excel.Range rngTitolo = _ws.Range[_ws.Cells[_rigaAttiva, colonnaInizio], _ws.Cells[_rigaAttiva, colonnaInizio + oreGiorno - 1]];

                _nomiDefiniti.Add(DefinedNames.GetName(entita["SiglaEntita"], "T", suffissoData), Tuple.Create(_rigaAttiva, colonnaInizio), Tuple.Create(_rigaAttiva, colonnaInizio + oreGiorno - 1));

                rngTitolo.Merge();
                rngTitolo.Style = "titleBarStyle";
                if (Struct.tipoVisualizzazione == "O")
                    rngTitolo.Value = entita["DesEntita"].ToString().ToUpperInvariant();
                else if (Struct.tipoVisualizzazione == "V")
                    rngTitolo.Value = giorno.ToString("MM/dd/yyyy");

                rngTitolo.RowHeight = 25;

                colonnaInizio += oreGiorno;
            });
            _rigaAttiva++;
        }
        protected virtual void InsertRangeGrafici(object siglaEntita, DateTime? giorno = null)
        {
            DataView grafici = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAGRAFICO].DefaultView;
            grafici.RowFilter = "SiglaEntita = '" + siglaEntita + "'";

            int i = 1;
            foreach (DataRowView grafico in grafici)
            {
                string suffissoData = giorno == null ? "DATA1" : Date.GetSuffissoData(giorno.Value);

                string graficoRange = DefinedNames.GetName(grafico["SiglaEntita"], "GRAFICO" + i++, suffissoData);

                Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, _colonnaInizio],
                    _ws.Cells[_rigaAttiva, _colonnaInizio + _intervalloOre - 1]];

                _nomiDefiniti.Add(graficoRange, Tuple.Create(_rigaAttiva, _colonnaInizio), Tuple.Create(_rigaAttiva, _colonnaInizio + _intervalloOre - 1));
                rng.Merge();
                rng.Style = "chartsBarStyle";
                rng.RowHeight = 200;
                _rigaAttiva++;
            }
        }
        protected virtual void InsertOre(object siglaEntita)
        {
            if (Struct.tipoVisualizzazione == "V")
            {
                int colonnaInizio = _colonnaInizio;
                
                InsertOre(_rigaAttiva, colonnaInizio, _dataInizio, 25);
                
                _rigaAttiva++;
            }
        }
        private void InsertOre(int rigaAttiva, int colonnaInizio, DateTime giorno, int oreGiorno)
        {
            Excel.Range rngOre = _ws.Range[_ws.Cells[rigaAttiva, colonnaInizio], _ws.Cells[rigaAttiva, colonnaInizio + oreGiorno - 1]];
            rngOre.Style = "dateBarStyle";
            rngOre.NumberFormat = "00";
            rngOre.Font.Size = 10;
            rngOre.RowHeight = 20;

            object[] valoriOre = new object[oreGiorno];
            for (int ora = 0; ora < oreGiorno; ora++)
            {
                int val = ora + 1;
                if (giorno == _dataInizio && _struttura.visData0H24)
                    val = ora == 0 ? 24 : ora;

                valoriOre[ora] = val;
            }
            rngOre.Value = valoriOre;
        }
        protected virtual void InsertTitoloVerticale(object siglaEntita, object siglaEntitaBreve, int numInformazioni)
        {
            int colonnaTitoloVert = _colonnaInizio - _visParametro - 1;
            Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, colonnaTitoloVert], _ws.Cells[_rigaAttiva + numInformazioni - 1, colonnaTitoloVert]];
            rng.Style = "titoloVertStyle";
            rng.Merge();

            _nomiDefiniti.Add(DefinedNames.GetName(siglaEntita, "TITOLO_VERTICALE", Date.GetSuffissoData(_dataInizio)), _rigaAttiva, colonnaTitoloVert, _rigaAttiva + numInformazioni - 1, colonnaTitoloVert);

            if (numInformazioni > 3) 
            {
                rng.Orientation = numInformazioni == 1 ? Excel.XlOrientation.xlHorizontal : Excel.XlOrientation.xlVertical;
                rng.Font.Size = numInformazioni == 1 ? 6 : 9;
                if (Struct.tipoVisualizzazione == "O")
                    rng.Value = siglaEntitaBreve;
                else if (Struct.tipoVisualizzazione == "V")
                {
                    rng.NumberFormat = "ddd d";
                    rng.Value = _dataInizio;
                }
            }
        }
        protected virtual void FormattaAllDati(object siglaEntita)
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            informazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
            informazioni.Sort = "Ordine";

            int rigaAttiva = _rigaAttiva;
            int rigaInizioGruppo = rigaAttiva;
            int allDatiIndice = 1;

            bool primaRigaTitolo2 = informazioni[0]["SiglaTipologiaInformazione"].ToString() == "TITOLO2";
            int ultimaColonna = 0;
            foreach (DataRowView info in informazioni)
            {
                bool isPrimaRiga = informazioni[0] == info;
                bool isUltimaRiga = informazioni[informazioni.Count - 1] == info;

                //se non è la prima riga, se è l'ultima, se è un inizio gruppo e se prima non ho sistemato un TITOLO2, creo un range ALLDATI
                if ((!isPrimaRiga && info["InizioGruppo"].ToString() == "1" && rigaInizioGruppo < rigaAttiva) || isUltimaRiga)
                {
                    int colonnaInizioAllDati = _colonnaInizio;
                    CicloGiorni((oreGiorno, suffissoData, giorno) =>
                    {
                        ultimaColonna = colonnaInizioAllDati + oreGiorno - 1;
                        int ultimaRiga = rigaAttiva - (isUltimaRiga ? 0 : 1);
                        Excel.Range allDati = _ws.Range[_ws.Cells[rigaInizioGruppo, colonnaInizioAllDati], _ws.Cells[ultimaRiga, ultimaColonna]];
                        allDati.Style = "allDatiStyle";
                        if (isUltimaRiga && rigaAttiva - rigaInizioGruppo == 1)
                            allDati.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlMedium;
                        allDati.EntireColumn.ColumnWidth = _cell.Width.dato;
                        allDati.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

                        if (Struct.tipoVisualizzazione == "V")
                        {
                            int deltaOre = 24 - Date.GetOreGiorno(giorno);
                            if (deltaOre >= 0)
                            {
                                Excel.Range rngOre = _ws.Range[_ws.Cells[rigaInizioGruppo, ultimaColonna - deltaOre], _ws.Cells[ultimaRiga, ultimaColonna]];
                                Style.RangeStyle(rngOre, "BackPattern:CrissCross");
                            }
                        }
                        colonnaInizioAllDati += oreGiorno;
                    });
                    ultimaColonna = colonnaInizioAllDati - 1;
                    allDatiIndice++;
                    rigaInizioGruppo = rigaAttiva + (info["SiglaTipologiaInformazione"].ToString() == "TITOLO2" ? 1 : 0);
                }
                if (isPrimaRiga && primaRigaTitolo2)
                    rigaInizioGruppo++;

                rigaAttiva++;
            }

            rigaAttiva = _rigaAttiva;
            foreach (DataRowView info in informazioni)
            {
                if (!info["SiglaTipologiaInformazione"].Equals("TITOLO2"))
                {
                    string grassetto = "Bold:" + info["Grassetto"];
                    string formato = "NumberFormat:[" + info["Formato"] + "]";
                    string align = "Align:" + Enum.Parse(typeof(Excel.XlHAlign), info["Align"].ToString());

                    Excel.Range rigaInfo = _ws.Range[_ws.Cells[rigaAttiva, _colonnaInizio], _ws.Cells[rigaAttiva, ultimaColonna]];
                    Style.RangeStyle(rigaInfo, grassetto + ";" + formato + ";" + align);
                }
                rigaAttiva++;
            }
        }
        protected virtual void InsertInformazioniEntita(object siglaEntita)
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            informazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
            informazioni.Sort = "Ordine";

            int rigaAttiva = _rigaAttiva;
            int colonnaTitoloInfo = _colonnaInizio - _visParametro;

            bool titolo2 = false;
            foreach (DataRowView info in informazioni)
            {
                string bordoTop = "Top:" + (informazioni[0] == info || (info["InizioGruppo"].ToString() == "1" && !titolo2) ? "medium" : "thin");
                string bordoBottom = "Bottom:" + (informazioni[informazioni.Count - 1] == info ? "medium" : "thin");
                int backColor = (info["BackColor"] is DBNull ? 0 : (int)info["BackColor"]);
                backColor = backColor == 0 || backColor == 2 ? (info["Editabile"].ToString() == "1" ? 15 : 48) : backColor;
                titolo2 = false;

                //proprietà di stile comuni
                string style = "FontSize:" + info["FontSize"] + ";FontName:Verdana;BackColor:" + backColor + ";"
                    + "ForeColor:" + info["ForeColor"] + ";Visible:" + info["Visibile"] + ";";

                //personalizzazioni a seconda della tipologia di informazione
                if (info["SiglaTipologiaInformazione"].Equals("TITOLO2"))
                {
                    Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaTitoloInfo], _ws.Cells[rigaAttiva, colonnaTitoloInfo + _intervalloOre + 1]];
                    style += "Bold:" + info["Grassetto"] + ";Merge:true;Borders:[" + bordoTop + ",Bottom:thin,Right:medium]";
                    Style.RangeStyle(rng, style);
                    rng.Value = info["DesInformazione"].ToString();
                    titolo2 = true;
                }
                else
                {
                    Excel.Range rng = _ws.Range[_ws.Cells[rigaAttiva, colonnaTitoloInfo], _ws.Cells[rigaAttiva, colonnaTitoloInfo + _visParametro - 1]];
                    style += "Borders:[insidev:thin,right:medium," + bordoTop + "," + bordoBottom + "]";
                    Style.RangeStyle(rng, style);

                    object[] valori = new object[_visParametro];
                    valori[0] = info["DesInformazione"];
                    valori[1] = info["DesInformazioneBreve"];

                    //TODO creare _struttura per COLONNA PARAMETRO                    
                    if (_struttura.visParametro)
                        valori[2] = "";

                    string nome = "";

                    if (!info["Selezione"].Equals("0"))
                    {
                        nome = (info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"]) + Simboli.UNION + "SEL" + info["Selezione"];
                    }

                    rng.Value = valori;
                    _ws.Cells[rigaAttiva, colonnaTitoloInfo + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    
                    if (info["Data0H24"].Equals("0") && _struttura.visData0H24)
                        Style.RangeStyle(_ws.Cells[rigaAttiva, _colonnaInizio], "BackPattern:CrissCross");
                }
                rigaAttiva++;
            }
            rigaAttiva++;
        }
        protected virtual void CreaNomiCelle(object siglaEntita)
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            informazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
            informazioni.Sort = "Ordine";

            int rigaAttiva = _rigaAttiva;
            foreach (DataRowView info in informazioni)
            {
                int oraAttiva = _colonnaInizio;
                siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    bool isVisibleData0H24 = giorno == _dataInizio && _struttura.visData0H24;
                    if (isVisibleData0H24)
                    {
                        _nomiDefiniti.Add(DefinedNames.GetName(siglaEntita, info["SiglaInformazione"], "DATA0", "H24"), rigaAttiva, oraAttiva++, info["Editabile"].Equals("1"), info["SalvaDB"].Equals("1"), info["AnnotaModifica"].Equals("1"));
                        oreGiorno--;
                    }

                    for (int i = 0; i < oreGiorno; i++)
                    {
                        _nomiDefiniti.Add(DefinedNames.GetName(siglaEntita, info["SiglaInformazione"], suffissoData, "H" + (i + 1)), rigaAttiva, oraAttiva++, info["Editabile"].Equals("1"), info["SalvaDB"].Equals("1"), info["AnnotaModifica"].Equals("1"));
                    }
                });
                rigaAttiva++;
            }
        }
        protected override void InsertPersonalizzazioni(object siglaEntita) { }
        protected virtual void InsertValoriCelle(object siglaEntita)
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            informazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND (ValoreDefault IS NOT NULL OR FormulaInCella = 1)";

            //carico tutti i dati reperibili durante la creazione del foglio

            foreach (DataRowView info in informazioni)
            {
                siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                Tuple<int,int>[] riga = _nomiDefiniti[info["DATA0H24"].Equals("0"), siglaEntita, info["SiglaInformazione"]];
                if (info["ValoreDefault"] != DBNull.Value)
                {
                    Excel.Range rng = _ws.Range[_nomiDefiniti.GetRange(riga)];
                    rng.Value = info["ValoreDefault"];
                }
                else if (info["FormulaInCella"].Equals("1"))
                {
                    int deltaNeg;
                    int deltaPos;
                    Stopwatch watch = Stopwatch.StartNew();
                    string formula = "=" + PreparaFormula(info, "DATA0", "DATA1", 24, out deltaNeg, out deltaPos);
                    watch.Stop();

                    if (info["SiglaTipologiaInformazione"].Equals("OTTIMO"))
                    {
                        Excel.Range optRng = _ws.Cells[riga[0].Item1, riga[0].Item2];
                        string rng = _nomiDefiniti.GetRange(riga[1], riga.Last());
                        optRng.Formula = "=SUM(" + rng + ")";
                        _ws.Range[rng].Formula = formula;
                    }
                    else
                    {
                        _ws.Range[_ws.Cells[riga.First().Item1, riga.First().Item2 + deltaNeg], _ws.Cells[riga.Last().Item1, riga.Last().Item2 - deltaPos]].Formula = formula;
                    }
                }
            }
        }
        protected virtual void InsertParametri(object siglaEntita)
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            informazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaTipologiaParametro IS NOT NULL";

            DataView parametriD = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPARAMETROD].DefaultView;
            DataView parametriH = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPARAMETROH].DefaultView;

            foreach (DataRowView info in informazioni)
            {
                siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    if (_nomiDefiniti.IsDefined(DefinedNames.GetName(siglaEntita, info["SiglaInformazione"])))
                    {
                        Tuple<int, int>[] riga = _nomiDefiniti[siglaEntita, info["SiglaInformazione"], suffissoData];
                        Excel.Range rng = _ws.Range[_nomiDefiniti.GetRange(riga)];
                    
                        parametriD.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaParametro = '" + info["SiglaTipologiaParametro"] + "' AND CONVERT(DataIV, System.Int32) <= " + giorno.ToString("yyyyMMdd") + " AND CONVERT(DataFV, System.Int32) >= " + giorno.ToString("yyyyMMdd");

                        if (parametriD.Count > 0)
                            rng.Value = parametriD[0]["Valore"];
                        else
                        {
                            parametriH.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaParametro = '" + info["SiglaTipologiaParametro"] + "' AND CONVERT(DataIV, System.Int32) <= " + giorno.ToString("yyyyMMdd") + " AND CONVERT(DataFV, System.Int32) >= " + giorno.ToString("yyyyMMdd");

                            parametriH.Sort = "Ora";

                            object[] values = parametriH.ToTable(false, "Valore").AsEnumerable().Select(r => r["Valore"]).ToArray();
                            rng.Value = values;
                        }
                    }
                    
                });
            }
        }
        protected virtual void CreaFormattazioneCondizionale(object siglaEntita)
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
            informazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
            
            DataView formattazione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONEFORMATTAZIONE].DefaultView;

            foreach (DataRowView info in informazioni)
            {
                siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                formattazione.RowFilter = (info["SiglaEntitaRif"] is DBNull ? "SiglaEntita" : "SiglaEntitaRif") + " = '" + siglaEntita + "' AND SiglaInformazione = '" + info["SiglaInformazione"] + "'";

                foreach (DataRowView format in formattazione)
                {
                    CicloGiorni((oreGiorno, suffissoData, giorno) =>
                    {
                        Tuple<int, int>[] riga = _nomiDefiniti[siglaEntita, info["SiglaInformazione"], suffissoData];
                        Excel.Range rng = _ws.Range[_nomiDefiniti.GetRange(riga)];

                        string[] valore = format["Valore"].ToString().Replace("\"", "").Split('|');
                        if (format["NomeCella"] != DBNull.Value)
                        {
                            Tuple<int,int> cella = _nomiDefiniti[siglaEntita, format["NomeCella"], suffissoData, "H1"][0];
                            string address = R1C1toA1(cella);

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
                            Excel.FormatCondition cond = rng.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Formula1: formula);

                            cond.Font.Color = format["ForeColor"];
                            cond.Font.Bold = format["Grassetto"].Equals("1");
                            if ((int)format["BackColor"] != 0)
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

                            Excel.FormatCondition cond = rng.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, format["Operatore"], formula1, formula2);

                            cond.Font.Color = format["ForeColor"];
                            cond.Font.Bold = format["Grassetto"].Equals("1");
                            if ((int)format["BackColor"] != 0)
                                cond.Interior.Color = format["BackColor"];

                            cond.Interior.Pattern = format["Pattern"];
                        }
                    });
                }
            }
        }

        #endregion

        protected void InsertGrafici()
        {
            DataView dvCE = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            DataView grafici = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAGRAFICO].DefaultView;
            DataView graficiInfo = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAGRAFICOINFORMAZIONE].DefaultView;

            dvCE.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "'";

            foreach (DataRowView entita in dvCE)
            {
                grafici.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
                graficiInfo.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";

                int i = 1;
                foreach (DataRowView grafico in grafici)
                {
                    string nome = DefinedNames.GetName(grafico["SiglaEntita"], "GRAFICO" + i++);
                    List<Tuple<int, int>[]> rangeGrafici = _nomiDefiniti.GetRanges(nome);

                    foreach (var rangeGrafico in rangeGrafici)
                    {
                        var cella = _ws.Cells[rangeGrafico[0].Item1, rangeGrafico[0].Item2];

                        var rigaGrafico = _ws.Range[_ws.Cells[rangeGrafico[0].Item1, rangeGrafico[0].Item2], _ws.Cells[rangeGrafico[1].Item1, rangeGrafico[1].Item2]];
                        Excel.Chart chart = _ws.ChartObjects().Add(rigaGrafico.Left, rigaGrafico.Top + 1, rigaGrafico.Width, rigaGrafico.Height - 2).Chart;

                        chart.Parent.Name = nome;

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
                            Tuple<int, int>[] rangeDati = _nomiDefiniti[grafico["SiglaEntita"], info["SiglaInformazione"]];
                            Excel.Range datiGrafico = _ws.Range[_nomiDefiniti.GetRange(rangeDati)];

                            var serie = chart.SeriesCollection().Add(datiGrafico);
                            serie.Name = info["DesInformazione"].ToString();
                            serie.ChartType = (Excel.XlChartType)info["ChartType"];
                            serie.Interior.ColorIndex = info["InteriorColor"];
                            serie.Border.ColorIndex = info["BorderColor"];
                            serie.Border.Weight = info["BorderWeight"];
                            serie.Border.LineStyle = info["BorderLineStyle"];
                        }

                        AggiornaGrafici(chart, rigaGrafico);
                    }
                }
            }
        }

        public override void CaricaInformazioni(bool all)
        {
            try
            {
                DataView dvCE = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
                DataView dvEP = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPROPRIETA].DefaultView;

                dvCE.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "'";
                _dataInizio = DataBase.DB.DataAttiva;

                //calcolo tutte le date e mantengo anche la data max
                DateTime dataFineMax = _dataInizio;
                Dictionary<object, DateTime> dateFineUP = new Dictionary<object, DateTime>();
                foreach (DataRowView entita in dvCE)
                {
                    dvEP.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                    if (dvEP.Count > 0)
                        dateFineUP.Add(entita["SiglaEntita"], _dataInizio.AddDays(double.Parse("" + dvEP[0]["Valore"])));
                    else
                        dateFineUP.Add(entita["SiglaEntita"], _dataInizio.AddDays(Struct.intervalloGiorni));

                    dataFineMax = new DateTime(Math.Max(dataFineMax.Ticks, dateFineUP[entita["SiglaEntita"]].Ticks));
                }

                DataView datiApplicazione = DataBase.DB.Select(DataBase.SP.APPLICAZIONE_INFORMAZIONE_H, "@SiglaCategoria=" + _siglaCategoria + ";@SiglaEntita=ALL;@DateFrom=" + _dataInizio.ToString("yyyyMMdd") + ";@DateTo=" + dataFineMax.ToString("yyyyMMdd") + ";@Tipo=1;@All=" + (all ? "1" : "0")).DefaultView;

                DataView insertManuali = new DataView();
                if (all)
                    insertManuali = DataBase.DB.Select(DataBase.SP.APPLICAZIONE_INFORMAZIONE_COMMENTO, "@SiglaCategoria=" + _siglaCategoria + ";@SiglaEntita=ALL;@DateFrom=" + _dataInizio.ToString("yyyyMMdd") + ";@DateTo=" + dataFineMax.ToString("yyyyMMdd") + ";@All=1").DefaultView;

                foreach (DataRowView entita in dvCE)
                {
                    datiApplicazione.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND CONVERT(Data, System.Int32) <= " + dateFineUP[entita["SiglaEntita"]].ToString("yyyyMMdd");
                    _dataFine = dateFineUP[entita["SiglaEntita"]];
                    CaricaInformazioniEntita(datiApplicazione);
                    if (all)
                    {
                        insertManuali.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND CONVERT(SUBSTRING(Data, 1, 8), System.Int32) <= " + dateFineUP[entita["SiglaEntita"]].ToString("yyyyMMdd");
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
        #region Informazioni

        private void CaricaInformazioniEntita(DataView datiApplicazione)
        {
            foreach (DataRowView dato in datiApplicazione)
            {
                DateTime dataDato = DateTime.ParseExact(dato["Data"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                //sono nel caso DATA0H24
                if(dataDato < DataBase.DataAttiva) 
                {
                    Tuple<int,int> cella = _nomiDefiniti[dato["SiglaEntita"], dato["SiglaInformazione"], "DATA0", "H24"][0];
                    _ws.Cells[cella.Item1, cella.Item2].Value = dato["H24"];
                } 
                else 
                {
                    Tuple<int, int>[] riga = _nomiDefiniti[dato["SiglaEntita"], dato["SiglaInformazione"], Date.GetSuffissoData(dataDato)];

                    List<object> o = new List<object>(dato.Row.ItemArray);
                    //elimino i campi inutili
                    o.RemoveRange(o.Count - 3, 3);
                    _ws.Range[_nomiDefiniti.GetRange(riga)].Value = o.ToArray();
                }
            }
        }
        private void CaricaCommentiEntita(DataView insertManuali)
        {
            foreach (DataRowView commento in insertManuali)
            {
                DateTime giorno = DateTime.ParseExact(commento["Data"].ToString().Substring(0, 8), "yyyyMMdd", CultureInfo.InvariantCulture);
                Tuple<int, int> cella = _nomiDefiniti[commento["SiglaEntita"], commento["SiglaInformazione"], Date.GetSuffissoData(giorno), Date.GetSuffissoOra(commento["Data"])][0];
                Excel.Range rng = _ws.Cells[cella.Item1, cella.Item2];
                rng.ClearComments();
                rng.AddComment("Valore inserito manualmente");
            }
        }

        #endregion

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
                        Tuple<int,int>[] riga;
                        if(all)
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

                    string nome = "";
                    if (parametroEntita.Length > 1)
                    {
                        int eRif = int.Parse(Regex.Match(parametroEntita[1], @"\d+").Value);
                        DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
                        categoriaEntita.RowFilter = "Gerarchia = '" + info["SiglaEntita"] + "' AND Riferimento = " + eRif;
                        nome = DefinedNames.GetName(categoriaEntita[0]["SiglaEntita"], parametri[n - 1]);
                    }
                    else
                        nome = DefinedNames.GetName(info["SiglaEntita"], parametri[n - 1]);

                    if (Regex.IsMatch(nome, @"\[[-+]?\d+\]"))
                    {
                        int deltaOre = int.Parse(nome.Split('[')[1].Replace("]", ""));

                        if (suffissoData == "DATA1")
                        {//traslo in avanti la formula di |deltaNeg| - |deltaOre|
                            int ora = Math.Abs(tmpdeltaNeg) + deltaOre + (info["Data0H24"].Equals("1") ? 0 : 1);
                            nome += Simboli.UNION + (ora == 0 ? DefinedNames.GetName("DATA0", "H24") : DefinedNames.GetName("DATA1", "H" + ora));
                        }
                        else
                        {
                            int ora = (deltaOre < 0 ? oreDataPrec + deltaOre + 1 : deltaOre + 1);
                            nome += Simboli.UNION + DefinedNames.GetName(deltaOre < 0 ? suffissoDataPrec : suffissoData, "H" + ora);
                        }
                        nome = Regex.Replace(nome, @"\[[-+]?\d+\]", "");
                    }
                    else
                    {
                        if (suffissoData == "DATA1")
                        {
                            int ora = tmpdeltaNeg == 0 ? 1 : Math.Abs(tmpdeltaNeg) + (info["Data0H24"].Equals("1") ? 0 : 1);
                            nome += Simboli.UNION + DefinedNames.GetName(suffissoData, "H" + ora);
                        }
                        else
                        {
                            nome += Simboli.UNION + DefinedNames.GetName(suffissoData, "H1");
                        }
                    }

                    Tuple<int, int> coordinate = _nomiDefiniti[nome][0];

                    return R1C1toA1(coordinate.Item1, coordinate.Item2);
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
            }
            CaricaInformazioni(all);
            AggiornaGrafici();
        }
        #region UpdateData

        //cancella i dati di tutti i giorni o del giorno specificato (attenzione ad usare in caso di cambio data perché il prefisso viene calcolato da data inizio config)
        private void CancellaDati(DateTime? giorno = null)
        {
            DataView dvCE = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            DataView dvEP = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPROPRIETA].DefaultView;

            dvCE.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "'"; // AND (Gerarchia = '' OR Gerarchia IS NULL )";

            string suffissoData = giorno == null ? null : Date.GetSuffissoData(DataBase.DB.DataAttiva, giorno.Value);

            foreach (DataRowView entita in dvCE)
            {
                DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAINFORMAZIONE].DefaultView;
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND FormulaInCella = '0' AND ValoreDefault IS NULL";

                foreach (DataRowView info in informazioni)
                {
                    if (Struct.tipoVisualizzazione == "O" || suffissoData != null)
                    {
                        var siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                        Tuple<int, int>[] riga = _nomiDefiniti[siglaEntita, info["SiglaInformazione"], suffissoData];

                        Excel.Range rng = _ws.Range[_nomiDefiniti.GetRange(riga)];
                        rng.Value = "";
                        rng.ClearComments();
                        Style.RangeStyle(rng, "BackColor:" + info["BackColor"] + ";ForeColor:" + info["ForeColor"]);
                    }
                    else if (Struct.tipoVisualizzazione == "V")
                    {
                        CicloGiorni(DataBase.DataAttiva, DataBase.DataAttiva.AddDays(Struct.intervalloGiorni), (oreGiorno, suffData, g) => 
                        {
                            var siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                            Tuple<int, int>[] riga = _nomiDefiniti[siglaEntita, info["SiglaInformazione"], suffData];

                            Excel.Range rng = _ws.Range[_nomiDefiniti.GetRange(riga)];
                            rng.Value = "";
                            rng.ClearComments();
                            Style.RangeStyle(rng, "BackColor:" + info["BackColor"] + ";ForeColor:" + info["ForeColor"]);
                        });
                    }
                    
                }
            }
        }
        public override void AggiornaDateTitoli()
        {
            DataView dvCE = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            DataView dvEP = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPROPRIETA].DefaultView;

            dvCE.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND Gerarchia IS NULL";
            _dataInizio = DataBase.DB.DataAttiva;

            foreach (DataRowView entita in dvCE)
            {
                dvEP.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                if (dvEP.Count > 0)
                    _dataFine = _dataInizio.AddDays(double.Parse("" + dvEP[0]["Valore"]));
                else
                    _dataFine = _dataInizio.AddDays(Struct.intervalloGiorni);

                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    if (Struct.tipoVisualizzazione == "O")
                    {
                        Tuple<int, int>[] range = _nomiDefiniti.GetRanges(DefinedNames.GetName(entita["SiglaEntita"], suffissoData))[0];
                        _ws.Range[_nomiDefiniti.GetRange(range)].Value = giorno;
                    }
                    else if (Struct.tipoVisualizzazione == "V")
                    {
                        Tuple<int, int>[] range = _nomiDefiniti.GetRanges(DefinedNames.GetName(entita["SiglaEntita"], "T", suffissoData))[0];
                        _ws.Range[_nomiDefiniti.GetRange(range)].Value = giorno;

                        range = _nomiDefiniti.GetRanges(DefinedNames.GetName(entita["SiglaEntita"], "TITOLO_VERTICALE", suffissoData))[0];
                        if(range[1].Item1 - range[0].Item1 > 3)
                            _ws.Range[_nomiDefiniti.GetRange(range)].Value = giorno;

                        range = _nomiDefiniti[DefinedNames.GetName(entita["SiglaEntita"], suffissoData, "GOTO")];
                        _ws.Cells[range[0].Item1, range[0].Item2].Value = giorno;
                    }
                    
                });
            }
        }

        protected void CaricaParametri()
        {
            DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPROPRIETA].DefaultView;
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;

            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";
            _dataInizio = DataBase.DB.DataAttiva;

            foreach (DataRowView entita in categoriaEntita)
            {
                entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                if (entitaProprieta.Count > 0)
                    _dataFine = _dataInizio.AddDays(double.Parse("" + entitaProprieta[0]["Valore"]));
                else
                    _dataFine = _dataInizio.AddDays(Struct.intervalloGiorni);

                InsertParametri(entita["SiglaEntita"]);
            }
        }

        public override void AggiornaGrafici()
        {
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIAENTITA].DefaultView;
            DataView grafici = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAGRAFICO].DefaultView;

            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "'";

            foreach (DataRowView entita in categoriaEntita)
            {
                grafici.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";                

                int i = 1;
                foreach (DataRowView grafico in grafici)
                {
                    string nome = DefinedNames.GetName(grafico["SiglaEntita"], "GRAFICO" + i++);

                    List<Tuple<int, int>[]> rangeGrafici = _nomiDefiniti.GetRanges(nome);

                    foreach (var rangeGrafico in rangeGrafici)
                    {
                        Excel.Range rigaGrafico = _ws.Range[_nomiDefiniti.GetRange(rangeGrafico)];
                        var chart = _ws.ChartObjects(nome).Chart;
                        AggiornaGrafici(chart, rigaGrafico);
                        chart.Refresh();
                    }
                }
            }
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
