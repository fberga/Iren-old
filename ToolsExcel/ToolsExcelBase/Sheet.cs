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
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Base
{
    /// <summary>
    /// Interfaccia con i metodi astratti o virtuali di creazione di un foglio contenente dati riferiti a impianti.
    /// </summary>
    public abstract class ASheet
    {
        #region Variabili

        protected Struct _struttura;
        protected DateTime _dataInizio;
        protected DateTime _dataFine;
        protected int _visParametro;

        protected static bool _protetto = true;

        #endregion

        #region Metodi

        /// <summary>
        /// In un ciclo che avanza di giorno in giorno da dataInizio a dataFine, esegui il delegato callback che definisce una routine specifica.
        /// </summary>
        /// <param name="dataInizio">Data di inizio del ciclo.</param>
        /// <param name="dataFine">Data di fine del ciclo.</param>
        /// <param name="callback">Delegato eseguito come corpo del ciclo.</param>
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
                callback(oreGiorno, suffissoData, giorno);
            }
        }
        /// <summary>
        /// In un ciclo che avanza di giorno in giorno a partire da DataBase.DataAttiva per il numero di giorni definito per l'entità, esegui il delegato callback che definisce una routine specifica.
        /// </summary>
        /// <param name="callback">Delegato eseguito come corpo del ciclo.</param>
        protected void CicloGiorni(Action<int, string, DateTime> callback)
        {
            CicloGiorni(_dataInizio, _dataFine, callback);
        }
        /// <summary>
        /// Metodo di caricamento della struttura del foglio.
        /// </summary>
        public abstract void LoadStructure();
        /// <summary>
        /// Metodo di aggiornamento dei dati del foglio.
        /// </summary>
        public abstract void UpdateData();
        /// <summary>
        /// Metodo di aggiornamento delle date dei titolo.
        /// </summary>
        public abstract void AggiornaDateTitoli();
        /// <summary>
        /// Metodi di aggiornamento dei grafici.
        /// </summary>
        public abstract void AggiornaGrafici();
        /// <summary>
        /// Metodo che permette di aggiungere delle customizzazioni durante la creazione della struttura.
        /// </summary>
        /// <param name="siglaEntita"></param>
        protected virtual void InsertPersonalizzazioni(object siglaEntita) { }
        /// <summary>
        /// Metodo per il caricamento delle informazioni.
        /// </summary>
        public abstract void CaricaInformazioni();
        
        #endregion

        #region Proprietà Statiche

        /// <summary>
        /// Restituisce o imposta la proprietà di protezione del workbook e dei fogli in esso contenuti.
        /// </summary>
        public static bool Protected
        {
            get { return _protetto; }
            set
            {
                if (_protetto != value)
                {
                    _protetto = value;

                    if (value)
                        Workbook.WB.Protect(Simboli.pwd);
                    else
                        Workbook.WB.Unprotect(Simboli.pwd);

                    foreach (Excel.Worksheet ws in Workbook.Sheets)
                    {
                        if (value)
                            if (ws.Name == "Log")
                                ws.Protect(Simboli.pwd, AllowSorting: true, AllowFiltering: true);
                            else
                                ws.Protect(Simboli.pwd);
                        else
                            ws.Unprotect(Simboli.pwd);
                    }
                }
            }
        }

        #endregion

        #region Metodi Statici

        /// <summary>
        /// Metodo per abilitare la modifica nelle informazioni per cui è concessa la modifica da DB.
        /// </summary>
        /// <param name="abilita">La modifica è abilitata se la proprietà è a true.</param>
        public static void AbilitaModifica(bool abilita)
        {
            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            categorie.RowFilter = "Operativa = '1'";

            Protected = false;
            foreach (DataRowView categoria in categorie)
            {
                Excel.Worksheet ws = Workbook.Sheets[categoria["DesCategoria"].ToString()];
                DefinedNames definedNames = new DefinedNames(categoria["DesCategoria"].ToString(), DefinedNames.InitType.Editable);

                foreach (string range in definedNames.Editable.Values)
                {
                    string[] subRanges = range.Split(',');
                    if (subRanges.Length == 1 && ws.Range[subRanges[0]].Cells.Count == 1)
                    {
                        ws.Range[subRanges[0]].Locked = !abilita;
                    }
                    else if (ws.Range[subRanges[0]].EntireRow.Hidden == false)
                    {
                        foreach (string subRange in subRanges)
                        {
                            ws.Range[subRange].Locked = !abilita;
                        }
                    }
                }
            }
            Protected = true;
        }
        /// <summary>
        /// Metodo che registra in DataBase.LocalDB le modifiche effettuate dall'utente durante il periodo in cui la modifica è attiva. Lavora con le sole entità modificate in modo da non sovraccaricare di modifiche il DB.
        /// </summary>
        public static void SalvaModifiche()
        {
            DataTable categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA];
            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            DataView entitaInformazione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;

            DataTable modifiche = DataBase.LocalDB.Tables[DataBase.Tab.MODIFICA];

            //controllo quali entità sono state modificate
            //List<object> entitaModificate =
            //    (from r in modifiche.AsEnumerable()
            //     group r["SiglaEntita"] by r["SiglaEntita"] into gr
            //     select gr.Key).ToList();

            foreach (DataRow entita in categoriaEntita.Rows)
            {
                object siglaEntita = entita["SiglaEntita"];
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                DefinedNames definedNames = new DefinedNames(nomeFoglio);

                Excel.Worksheet ws = Workbook.Sheets[nomeFoglio];

                bool hasData0H24 = definedNames.HasData0H24;

                entitaInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND ((FormulaInCella = '1' AND WB = '0' AND SalvaDB = '1') OR (WB <> '0' AND SalvaDB = '1'))";

                DataTable entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA];
                DateTime dataFine = DataBase.DataAttiva.AddDays(Math.Max(
                    (from r in entitaProprieta.AsEnumerable()
                     where r["SiglaEntita"].Equals(siglaEntita) && r["SiglaProprieta"].ToString().EndsWith("GIORNI_STRUTTURA")
                     select int.Parse(r["Valore"].ToString())).FirstOrDefault(), Struct.intervalloGiorni));

                foreach (DataRowView info in entitaInformazione)
                {
                    object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                    DateTime giorno = DataBase.DataAttiva;

                    //prima cella della riga da salvare (non considera Data0H24)
                    Range rng = definedNames.Get(siglaEntitaRif, info["SiglaInformazione"], Date.SuffissoDATA1).Extend(colOffset: Date.GetOreIntervallo(dataFine));

                    Handler.StoreEdit(ws.Range[rng.ToString()], 0, true);
                }
            }
        }        

        #endregion
    }
    /// <summary>
    /// Classe base con i metodi per la creazione di un foglio contenente dati riferiti a impianti.
    /// </summary>
    public class Sheet : ASheet, IDisposable
    {
        #region Variabili

        protected Excel.Worksheet _ws;
        protected object _siglaCategoria;
        protected DefinedNames _definedNames;
        protected int _intervalloOre;
        protected int _rigaAttiva;
        protected bool _disposed = false;
        protected int _intervalloGiorniMax;
        protected Dictionary<object, DateTime> _dataFineUP = new Dictionary<object, DateTime>();

        #endregion

        #region Costruttori

        public Sheet(Excel.Worksheet ws)
        {
            _ws = ws;

            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            categorie.RowFilter = "DesCategoria = '" + ws.Name + "'";

            _siglaCategoria = categorie[0]["SiglaCategoria"];

            AggiornaParametriSheet();
            _definedNames = new DefinedNames(_ws.Name);

            //carico la massima datafine in maniera da creare la barra navigazione della dimensione giusta (compresa la definizione dei giorni se necessario)
            DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "'";

            foreach (DataRowView entita in categoriaEntita)
            {
                entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta LIKE '%GIORNI_STRUTTURA'";
                int intervalloGiorni = entitaProprieta.Count > 0 ? int.Parse(entitaProprieta[0]["Valore"].ToString()) : Struct.intervalloGiorni;

                _dataFineUP.Add(entita["SiglaEntita"], DataBase.DataAttiva.AddDays(intervalloGiorni));
                _intervalloGiorniMax = Math.Max(_intervalloGiorniMax, intervalloGiorni);
            }
        }
        ~Sheet()
        {
            Dispose();
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Legge da LocalDB i parametri dell'applicazione e reinizializza tutte le strutture del workbook e del foglio.
        /// </summary>
        protected void AggiornaParametriSheet()
        {
            DataView paramApplicazione = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE].DefaultView;

            _struttura = new Struct();

            //cerco selezioni
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "'";

            DataView entitaInformazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
            bool visSelezione = false;
            foreach (DataRowView entita in categoriaEntita)
            {
                entitaInformazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND Selezione > 0";
                if(entitaInformazioni.Count > 0)
                {
                    visSelezione = true;
                    break;
                }
            }

            _struttura.rigaBlock = (int)paramApplicazione[0]["RowBlocco"];
            _struttura.rigaGoto = (int)paramApplicazione[0]["RowGoto"];
            _struttura.visData0H24 = paramApplicazione[0]["VisData0H24"].ToString() == "1";
            _struttura.visParametro = paramApplicazione[0]["VisParametro"].ToString() == "1";
            _struttura.visSelezione = visSelezione;
            _struttura.colBlock = (int)paramApplicazione[0]["ColBlocco"] + (_struttura.visParametro ? 1 : 0) + (visSelezione ? 1 : 0);
            Struct.tipoVisualizzazione = paramApplicazione[0]["TipoVisualizzazione"] is DBNull ? "O" : paramApplicazione[0]["TipoVisualizzazione"].ToString();
            Struct.intervalloGiorni = paramApplicazione[0]["IntervalloGiorniEntita"] is DBNull ? 0 : (int)paramApplicazione[0]["IntervalloGiorniEntita"];
            Struct.visualizzaRiepilogo = paramApplicazione[0]["VisRiepilogo"] is DBNull ? true : paramApplicazione[0]["VisRiepilogo"].Equals("1");

            _visParametro = _struttura.visParametro ? 3 : 2 + (visSelezione ? 1 : 0);

            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL)";
            _struttura.numEleMenu = (Struct.tipoVisualizzazione == "O" ? categoriaEntita.Count : (Struct.intervalloGiorni + 1));
            _struttura.numRigheMenu = 1;
            if (_struttura.numEleMenu > 8)
            {
                int tmp = _struttura.numEleMenu;
                while (tmp / 8 > 0)
                {
                    _struttura.rigaBlock++;
                    _struttura.numRigheMenu++;
                    tmp /= 8;
                }
            }
        }
        /// <summary>
        /// Launcher per l'aggiornamento della struttura. Definisce anche le colonne in base all'intervallo massimo di giorni delle entità presenti nel foglio.
        /// </summary>
        public override void LoadStructure()
        {
            SplashScreen.UpdateStatus("Aggiorno struttura " + _ws.Name);

            DataTable entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA];
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;

            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL)";
            
            _dataInizio = Utility.DataBase.DB.DataAttiva;
            _dataFine = Utility.DataBase.DB.DataAttiva.AddDays(Struct.tipoVisualizzazione == "O" ? _intervalloGiorniMax : 0);

            //Definizione dei nomi delle colonne
            _definedNames.DefineDates(_dataInizio, _dataFine, _struttura.colBlock, _struttura.visData0H24);

            Stopwatch watch = Stopwatch.StartNew();
            Clear();
            watch.Stop();
            watch = Stopwatch.StartNew();
            
            InitBarraNavigazione();

            watch.Stop();
            

            _rigaAttiva = _struttura.rigaBlock + 1;

            foreach (DataRowView entita in categoriaEntita)
            {
                string siglaEntita = "" + entita["SiglaEntita"];

                if (Struct.tipoVisualizzazione == "O")
                {
                    _dataFine = _dataInizio.AddDays(Math.Max(
                        (from r in entitaProprieta.AsEnumerable()
                         where r["SiglaEntita"].Equals(siglaEntita) && r["SiglaProprieta"].ToString().EndsWith("GIORNI_STRUTTURA")
                         select int.Parse(r["Valore"].ToString())).FirstOrDefault(), Struct.intervalloGiorni));

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

            //entitaProprieta.RowFilter = "";
            categoriaEntita.RowFilter = "";

            _definedNames.DumpToDataSet();
            CaricaInformazioni();
            AggiornaGrafici();

            //CalcolaFormule();                     //TODO

            //cancello tutte le selezioni
            //_ws.Activate();
            //_ws.Cells[1, 1].Select();
            //Workbook.Main.Select();
            //Workbook.ScreenUpdating = false;
        }
        /// <summary>
        /// Metodo per eliminare la struttura esistente dal foglio e prepararlo alla nuova che verrà caricata.
        /// </summary>
        protected void Clear()
        {
            SplashScreen.UpdateStatus("Cancello struttura foglio '" + _ws.Name + "'");

            _ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            if (_ws.ChartObjects().Count > 0)
                _ws.ChartObjects().Delete();

            _ws.Rows.ClearContents();
            _ws.Rows.ClearComments();
            _ws.Rows.FormatConditions.Delete();
            _ws.Rows.EntireRow.Hidden = false;
            _ws.Rows.Style = "Normal";

            _ws.Rows.RowHeight = Struct.cell.height.normal;
            _ws.Columns.ColumnWidth = Struct.cell.width.dato;

            _ws.Rows["1:" + (_struttura.rigaBlock - 1)].RowHeight = Struct.cell.height.empty;

            for (int i = 0; i < _struttura.numRigheMenu; i++)
                _ws.Rows[_struttura.rigaGoto + i].RowHeight = Struct.cell.height.normal;

            _ws.Columns[1].ColumnWidth = Struct.cell.width.empty;
            _ws.Columns[2].ColumnWidth = Struct.cell.width.entita;

            if (!Aggiorna._freezePanes.ContainsKey(_ws.Name) ||
                (Aggiorna._freezePanes[_ws.Name].Item1 != _struttura.rigaBlock || Aggiorna._freezePanes[_ws.Name].Item2 != _struttura.colBlock))
            {
                ((Excel._Worksheet)_ws).Activate();
                _ws.Application.ActiveWindow.FreezePanes = false;
                _ws.Cells[_struttura.rigaBlock, _struttura.colBlock].Select();
                //_ws.Application.ActiveWindow.ScrollColumn = 1;
                //_ws.Application.ActiveWindow.ScrollRow = 1;
                _ws.Application.ActiveWindow.FreezePanes = true;
                Workbook.Main.Activate();
                //Workbook.ScreenUpdating = false;
            }
            

            int colInfo = _struttura.colBlock - _visParametro;
            _ws.Columns[colInfo].ColumnWidth = Struct.cell.width.informazione;
            _ws.Columns[colInfo + 1].ColumnWidth = Struct.cell.width.unitaMisura;
            if (_struttura.visSelezione)
                _ws.Columns[colInfo + 2].ColumnWidth = 2.5;
            if (_struttura.visParametro)
                _ws.Columns[colInfo + _visParametro].ColumnWidth = Struct.cell.width.parametro;
        }
        /// <summary>
        /// Inizializza la barra di navigazione nella parte alta del foglio applicandovi lo stile "Top menu GOTO". Definisce tutte le celle e genera gli oggetti GOTO per i tasti del menù e applica lo stile "Barra navigazione con nomi" se la visualizzazione è Orizzontale altrimenti "Barra navigazione con date". Se la tipologia di visualizzazione è Orizzontale, aggiunge anche la barra della data e delle ore (applicando "Barra della data").
        /// </summary>
        protected void InitBarraNavigazione()
        {
            SplashScreen.UpdateStatus("Inizializzo barra di navigazione '" + _ws.Name + "'");

            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";

            int dataOreTot = (Struct.tipoVisualizzazione == "O" ? Date.GetOreIntervallo(_dataInizio, _dataFine) : 25) + (_struttura.visData0H24 ? 1 : 0) + (_struttura.visParametro ? 1 : 0);
                
            Excel.Range gotoBar = _ws.Range[_ws.Cells[2, 2], _ws.Cells[_struttura.rigaGoto + _struttura.numRigheMenu, _struttura.colBlock + dataOreTot - 1]];
            gotoBar.Style = "Top menu GOTO";
            gotoBar.BorderAround2(Weight: Excel.XlBorderWeight.xlMedium, Color: 1);

            //scrivo nome applicazione in alto a sinistra
            Range title = new Range(_struttura.rigaGoto, 2, _struttura.numRigheMenu, _struttura.colBlock - 2);

            int fontSize = 12;
            double rangeSize = _ws.Range[title.ToString()].Width;
            for (; fontSize > 0; fontSize--)
            {
                Graphics grfx = Graphics.FromImage(new Bitmap(1, 1));
                grfx.PageUnit = GraphicsUnit.Point;
                SizeF sizeMax = grfx.MeasureString(Simboli.nomeApplicazione.ToUpper(), new Font("Verdana", fontSize, FontStyle.Bold));
                if (rangeSize > sizeMax.Width)
                    break;
            }

            Style.RangeStyle(_ws.Range[title.ToString()], merge: true, bold: true, fontSize: fontSize, align: Excel.XlHAlign.xlHAlignCenter);
            _ws.Range[title.ToString()].Value = Simboli.nomeApplicazione.ToUpper();

            //calcolo numero elementi per riga
            double numEleRiga = _struttura.numEleMenu / Convert.ToDouble(_struttura.numRigheMenu);

            int j = 0;
            for (int i = 0; i < _struttura.numEleMenu; i++)
            {
                int r = (i / (int)Math.Ceiling(numEleRiga));
                int c = (i % (int)Math.Ceiling(numEleRiga));

                object nome = Struct.tipoVisualizzazione == "O" ? categoriaEntita[i]["SiglaEntita"] : DefinedNames.GetName(categoriaEntita[0]["SiglaEntita"], Date.GetSuffissoData(DataBase.DataAttiva.AddDays(i)));

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
                
                _definedNames.AddGOTO(nome, Range.R1C1toA1(_struttura.rigaGoto + r, _struttura.colBlock + c + (_struttura.visData0H24 ? 1 : 0)));
                
                rng.Value = Struct.tipoVisualizzazione == "O" ? categoriaEntita[i]["DesEntitaBreve"] : DataBase.DataAttiva.AddDays(i);
                rng.Style = Struct.tipoVisualizzazione == "O" ? "Barra navigazione con nomi" : "Barra navigazione con date";
            }

            //inserisco la data e le ore
            if (Struct.tipoVisualizzazione == "O")
            {
                int colonnaInizio = _struttura.colBlock;
                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    int escludiH24 = (giorno == _dataInizio && _struttura.visData0H24 ? 1 : 0);


                    Range rngData = new Range(_struttura.rigaBlock - 2, colonnaInizio + escludiH24, 1, oreGiorno);

                    Excel.Range rng = _ws.Range[rngData.ToString()];
                    rng.Merge();
                    rng.Style = "Barra della data";
                    rng.Value = giorno.ToString("MM/dd/yyyy");
                    rng.RowHeight = 25;

                    Range rngOre = new Range(_struttura.rigaBlock - 1, colonnaInizio, 1, oreGiorno + escludiH24);
                    InsertOre(rngOre, giorno == _dataInizio && _struttura.visData0H24);
                    colonnaInizio += oreGiorno + escludiH24;
                });
            }
        }
        /// <summary>
        /// Launcher per le azioni di creazione del blocco entità.
        /// </summary>
        /// <param name="entita">La riga contenente le informazioni dell'entità.</param>
        protected void InitBloccoEntita(DataRowView entita)
        {
            Stopwatch watch = Stopwatch.StartNew();
            SplashScreen.UpdateStatus("Carico struttura " + entita["DesEntita"]);

            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
            DataView grafici = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_GRAFICO].DefaultView;
            DataView graficiInfo = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_GRAFICO_INFORMAZIONE].DefaultView;

            if (informazioni.RowFilter != "SiglaEntita = '" + entita["SiglaEntita"] + "'")
            {
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
                //informazioni.Sort = "Ordine";
                grafici.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
                graficiInfo.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
            }
            
            _intervalloOre = Date.GetOreIntervallo(_dataInizio, _dataFine) + (_struttura.visData0H24 ? 1 : 0) + (_struttura.visParametro ? 1 : 0);

            Stopwatch watch1 = Stopwatch.StartNew();
            CreaNomiCelle(entita["SiglaEntita"]);
            InsertTitoloEntita(entita["SiglaEntita"], entita["DesEntita"]);
            InsertOre(entita["SiglaEntita"]);
            InsertTitoloVerticale(entita["DesEntitaBreve"]);
            FormattaBloccoEntita();
            InsertInformazioniEntita();
            InsertPersonalizzazioni(entita["SiglaEntita"]);
            InsertGrafici();
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND (ValoreDefault IS NOT NULL OR FormulaInCella = 1)";
            InsertFormuleValoriDefault();
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaTipologiaParametro IS NOT NULL";
            InsertParametri();
            watch1.Stop();
            Stopwatch watch2 = Stopwatch.StartNew();
            informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "'";
            watch2.Stop();
            Stopwatch watch3 = Stopwatch.StartNew();
            FormattazioneCondizionale();
            watch3.Stop();

            watch.Stop();
            //due righe vuote tra un'entità e la successiva
            _rigaAttiva += 2;
        }
        #region Blocco entità

        /// <summary>
        /// Crea i nomi delle celle in base alla riga. Definisce se sono editabili o meno, se sono parte di selezioni, se devono essere salvate sul DB e se alla modifica deve essere segnalata la modifica stessa o meno. Collega i GOTO generati nel menu con la posizione effettiva delle entità.
        /// </summary>
        /// <param name="siglaEntita">La sigla dell'entità di cui creare i nomi.</param>
        protected virtual void CreaNomiCelle(object siglaEntita)
        {
            //inserisco titoli
            string suffissoData = Date.GetSuffissoData(_dataInizio);
            _definedNames.AddName(_rigaAttiva, Struct.tipoVisualizzazione == "O" ? siglaEntita : suffissoData, "T");
            //_definedNames.AddName(_rigaAttiva, siglaEntita, "T", Struct.tipoVisualizzazione == "O" ? "" : suffissoData);

            //sistemo l'indirizzamento dei GOTO
            int col = _definedNames.GetColFromDate(suffissoData);
            object name = Struct.tipoVisualizzazione == "O" ? siglaEntita : DefinedNames.GetName(siglaEntita, suffissoData);
            _definedNames.ChangeGOTOAddressTo(name, Range.R1C1toA1(_rigaAttiva, col));

            //aggiungo la riga delle ore
            _rigaAttiva += Struct.tipoVisualizzazione == "V" ? 2 : 1;

            //aggiungo i grafici
            DataView grafici = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_GRAFICO].DefaultView;

            int i = 1;
            foreach (DataRowView grafico in grafici)
            {
                _definedNames.AddName(_rigaAttiva, grafico["SiglaEntita"], "GRAFICO" + i, Struct.tipoVisualizzazione == "O" ? "" : Date.GetSuffissoData(_dataInizio));
                i++;
                _rigaAttiva++;
            }

            //aggiungo informazioni
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
            //_definedNames.AddName(_rigaAttiva, Struct.tipoVisualizzazione == "O" ? siglaEntita : suffissoData, "TITOLO_VERTICALE");

            int startCol = _definedNames.GetFirstCol();
            int colOffsett = _definedNames.GetColOffset();
            int remove25hour = (Struct.tipoVisualizzazione == "O" ? 0 : 25 - Date.GetOreGiorno(_dataInizio));
            bool isSelection = false;
            string rifSel = "";
            Dictionary<string, int> peers = new Dictionary<string, int>();

            foreach (DataRowView info in informazioni)
            {
                object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                _definedNames.AddName(_rigaAttiva, siglaEntitaRif, info["SiglaInformazione"], Struct.tipoVisualizzazione == "O" ? "" : Date.GetSuffissoData(_dataInizio));

                int data0H24 = (info["Data0H24"].Equals("0") && _struttura.visData0H24 ? 1 : 0);

                //selezione - Mantenere in questo ordine: alla prima volta entra nel selezione = 10, poi in isSelection e alla fine chiude la selezione e salta gli altri (se non in presenza di un altro 10)
                if (isSelection && (info["Selezione"].Equals(0) || info["Selezione"].Equals(10)))
                {
                    //salvo la selezione
                    _definedNames.SetSelection(rifSel, peers);
                    //chiudo selezione
                    isSelection = false;
                    rifSel = "";
                    peers = new Dictionary<string, int>();
                }
                if (isSelection)
                {
                    Range rng = new Range(_rigaAttiva, startCol - 1);
                    peers.Add(rng.ToString(), int.Parse(info["Selezione"].ToString()));
                }
                if (info["Selezione"].Equals(10))
                {
                    Range rng = new Range(_rigaAttiva, startCol + data0H24, 1, _definedNames.GetColOffset(_dataFine) - data0H24 - remove25hour);
                    isSelection = true;
                    rifSel = rng.ToString();
                }
                //fine selezione
                
                if (info["Editabile"].Equals("1"))
                {
                    
                    if (info["SiglaTipologiaInformazione"].Equals("GIORNALIERA"))
                    {
                        //seleziono la cella dell'unità di misura
                        Range rng = new Range(_rigaAttiva, startCol - _visParametro + 1);
                        _definedNames.SetEditable(_rigaAttiva, rng);
                    }
                    else
                    {
                        Range rng = new Range(_rigaAttiva, startCol + data0H24, 1, _definedNames.GetColOffset(_dataFine) - data0H24 - remove25hour);
                        _definedNames.SetEditable(_rigaAttiva, rng);
                    }
                }
                if (info["SalvaDB"].Equals("1"))
                    _definedNames.SetSaveDB(_rigaAttiva);

                if (info["AnnotaModifica"].Equals("1"))
                    _definedNames.SetToNote(_rigaAttiva);

                if (info["SiglaTipologiaInformazione"].Equals("CHECK") && info["Funzione"] != DBNull.Value)
                {
                    int checkType = int.Parse(Regex.Match(info["Funzione"].ToString(), @"\d+").Value);
                    Range rng = new Range(_rigaAttiva, startCol + data0H24, 1, colOffsett - data0H24 - remove25hour);
                    _definedNames.AddCheck(siglaEntitaRif.ToString(), rng.ToString(), checkType);
                }

                _rigaAttiva++;
            }
        }
        /// <summary>
        /// Applica lo stile "Barra titiolo entita" alla riga del titolo dell'entità e scrive la descrizione.
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entità per la ricerca della riga.</param>
        /// <param name="desEntita">Descrizione da scrivere nella riga.</param>
        protected virtual void InsertTitoloEntita(object siglaEntita, object desEntita)
        {
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                Range rng = Struct.tipoVisualizzazione == "O" ? _definedNames.Get(siglaEntita, "T", suffissoData) : _definedNames.Get(suffissoData, "T");
                rng.Extend(1, oreGiorno);

                Excel.Range rngTitolo = _ws.Range[rng.ToString()];
                rngTitolo.Merge();
                rngTitolo.Style = "Barra titolo entita";
                rngTitolo.Value = Struct.tipoVisualizzazione == "O" ? desEntita.ToString().ToUpperInvariant() : giorno.ToString("MM/dd/yyyy");
                rngTitolo.RowHeight = 25;
            });
        }
        /// <summary>
        /// In caso di visualizzazione Verticale, formatta e compila la barra delle ore.
        /// </summary>
        /// <param name="siglaEntita">Sigla dell'entità per la ricerca della riga.</param>
        protected virtual void InsertOre(object siglaEntita)
        {
            if (Struct.tipoVisualizzazione == "V")
            {
                Range rng = _definedNames.Get(Date.GetSuffissoData(_dataInizio), "T");
                rng.StartRow++;
                rng.Extend(1, 25);
                InsertOre(rng);
            }
        }
        /// <summary>
        /// Inserisce le ore nel range rng passato per parametro. Se hasData0H24 è true, mette nella prima cella 24.
        /// </summary>
        /// <param name="rng">Range su cui scrivere le ore</param>
        /// <param name="hasData0H24">True se è presente l'ora 24 del giorno precedente.</param>
        private void InsertOre(Range rng, bool hasData0H24 = false)
        {
            Excel.Range rngOre = _ws.Range[rng.ToString()];
            rngOre.Style = "Barra della data";
            rngOre.NumberFormat = "0";
            rngOre.Font.Size = 10;
            rngOre.RowHeight = 20;

            int ora = 1;
            foreach (Range cell in rng.Columns)
            {
                if (hasData0H24) 
                {
                    _ws.Range[cell.ToString()].Value = 24;
                    hasData0H24 = false;
                }
                else
                    _ws.Range[cell.ToString()].Value = ora++;
            }
        }
        /// <summary>
        /// Applica lo stile "Barra titolo verticale" con alcune modifiche e scrive la descrizione dell'entità.
        /// </summary>
        /// <param name="desEntita">Descrizione entità da scrivere.</param>
        protected virtual void InsertTitoloVerticale(object desEntita)
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;

            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            Range rngTitolo = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), _struttura.colBlock - _visParametro - 1, informazioni.Count);
            
            Stopwatch watch = Stopwatch.StartNew();
            Excel.Range titoloVert = _ws.Range[rngTitolo.ToString()];
            Style.RangeStyle(titoloVert, style: "Barra titolo verticale", orientation: informazioni.Count == 1 ? Excel.XlOrientation.xlHorizontal : Excel.XlOrientation.xlVertical, merge: true, fontSize: informazioni.Count == 1 ? 6 : 9, numberFormat: informazioni.Count > 4 ? "ddd d" : "dd");
            watch.Stop();
            titoloVert.Value = Struct.tipoVisualizzazione == "O" ? desEntita : _dataInizio;            
        }
        /// <summary>
        /// Formatta l'area dati impostando lo stile di base per le informazioni ("Area dati") e imposta, se ci sono, gli spazi per le informazioni giornaliere.
        /// </summary>
        protected virtual void FormattaBloccoEntita()
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;

            informazioni.RowFilter += " AND SiglaTipologiaInformazione <> 'GIORNALIERA'";

            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            Range rng = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), _definedNames.GetFirstCol() - _visParametro, informazioni.Count, _definedNames.GetColOffset(_dataFine) + _visParametro);

            Excel.Range bloccoEntita = _ws.Range[rng.ToString()];
            bloccoEntita.Style = "Area dati";
            bloccoEntita.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            bloccoEntita.Columns[1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            bloccoEntita.Columns[2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            bloccoEntita.Columns[_visParametro].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
            if (_struttura.visSelezione)
                bloccoEntita.Columns[3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            if (_struttura.visParametro)
                bloccoEntita.Columns[3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            int col = _struttura.visData0H24 ? 1 : 0;
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                bloccoEntita.Columns[_visParametro + col].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
                col += oreGiorno;
            });

            informazioni.RowFilter = informazioni.RowFilter.Replace(" AND SiglaTipologiaInformazione <> 'GIORNALIERA'", " AND SiglaTipologiaInformazione = 'GIORNALIERA'");
            if (informazioni.Count > 0)
            {
                rng = new Range(rng.StartRow + rng.RowOffset, rng.StartColumn, informazioni.Count, 2);
                bloccoEntita = _ws.Range[rng.ToString()];
                bloccoEntita.Style = "Area dati";
                bloccoEntita.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
                bloccoEntita.Columns[1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                bloccoEntita.Columns[2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            informazioni.RowFilter = informazioni.RowFilter.Replace(" AND SiglaTipologiaInformazione = 'GIORNALIERA'", "");
        }
        /// <summary>
        /// Inserisce le informazioni e applica la formattazione riga per riga in base alle informazioni sul DB.
        /// </summary>
        protected virtual void InsertInformazioniEntita()
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
            int col = _definedNames.GetFirstCol();
            int colOffset = _definedNames.GetColOffset(_dataFine);
            object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];
            int row = _definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], Date.GetSuffissoData(_dataInizio));

            Excel.Range rngRow = _ws.Range[Range.GetRange(row, col - _visParametro, informazioni.Count, colOffset + _visParametro)];
            Excel.Range rngInfo = _ws.Range[Range.GetRange(row, col - _visParametro, informazioni.Count, 2)];
            Excel.Range rngData = _ws.Range[Range.GetRange(row, col, informazioni.Count, colOffset)];

            if(Struct.tipoVisualizzazione == "V")
            {
                DataView infoNoGiornaliere = new DataView(DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE]);
                infoNoGiornaliere.RowFilter = informazioni.RowFilter + " AND SiglaTipologiaInformazione <> 'GIORNALIERA'";

                Excel.Range rngDataNoGiornaliere = _ws.Range[Range.GetRange(row, col, infoNoGiornaliere.Count, colOffset)];

                int oreGiorno = Date.GetOreGiorno(_dataInizio);
                if(oreGiorno < 24)
                    rngDataNoGiornaliere.Columns[rngDataNoGiornaliere.Columns.Count - 1].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                if(oreGiorno < 25)
                    rngDataNoGiornaliere.Columns[rngDataNoGiornaliere.Columns.Count].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
            }

            int i = 1;
            foreach (DataRowView info in informazioni)
            {
                rngInfo.Rows[i].Value = new object[2] { info["DesInformazione"], info["DesInformazioneBreve"] };

                int infoBackColor = info["Editabile"].ToString() == "1" ? 15 : 48;

                if(info["Selezione"].Equals(0) && _struttura.visSelezione)
                    rngRow.Rows[i].Cells[_visParametro].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;

                if (info["SiglaTipologiaInformazione"].Equals("GIORNALIERA"))
                {
                    Style.RangeStyle(rngInfo.Rows[i].Cells[1], 
                        fontSize: info["FontSize"], 
                        foreColor: info["ForeColor"],
                        backColor: (info["Editabile"].ToString() == "1" ? 15 : 48), 
                        visible: info["Visibile"].Equals("1"));

                    Style.RangeStyle(rngInfo.Rows[i].Cells[2], 
                        fontSize: info["FontSize"],
                        foreColor: info["ForeColor"],
                        backColor: info["BackColor"],
                        bold: info["Grassetto"].Equals("1"),
                        numberFormat: info["Formato"],
                        align: Enum.Parse(typeof(Excel.XlHAlign), info["Align"].ToString()));
                }
                else if (info["SiglaTipologiaInformazione"].Equals("TITOLO2"))
                {
                    Style.RangeStyle(rngRow.Rows[i], 
                        fontSize: info["FontSize"],
                        foreColor: info["ForeColor"],
                        backColor: info["BackColor"],
                        merge: true,
                        bold:true,
                        borders: "[Top:medium, Right:medium]");
                }
                else 
                {
                    if (info["InizioGruppo"].Equals("1"))
                        rngRow.Rows[i].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;

                    Style.RangeStyle(rngInfo.Rows[i], 
                        fontSize: info["FontSize"],
                        foreColor: info["ForeColor"],
                        backColor: infoBackColor,
                        visible: info["Visibile"].Equals("1"),
                        borders: "[Right:medium]");

                    Style.RangeStyle(rngData.Rows[i], 
                        fontSize: info["FontSize"],
                        foreColor: info["ForeColor"],
                        backColor: info["BackColor"],
                        bold: info["Grassetto"].Equals("1"),
                        numberFormat: info["Formato"],
                        align: Enum.Parse(typeof(Excel.XlHAlign), info["Align"].ToString()));

                    if (info["Data0H24"].Equals("0") && _struttura.visData0H24 && !info["SiglaTipologiaInformazione"].Equals("GIORNALIERA"))
                        rngData.Rows[i].Cells[1].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                }
                i++;
            }
        }
        /// <summary>
        /// Inserisce i valori di default e le formule.
        /// </summary>
        protected virtual void InsertFormuleValoriDefault()
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
            int colOffset = _definedNames.GetColOffset(_dataFine);
            foreach (DataRowView info in informazioni)
            {
                object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                //tolgo la colonna della DATA0H24 dove non serve
                int offsetAdjust = (_struttura.visData0H24 && info["Data0H24"].Equals("0") ? 1 : 0);

                Range rng = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, info["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), _definedNames.GetFirstCol());

                if (info["SiglaTipologiaInformazione"].Equals("GIORNALIERA"))
                    rng.StartColumn -= _visParametro - 1;
                else
                {
                    rng.StartColumn += offsetAdjust;
                    rng.Extend(colOffset: colOffset - offsetAdjust);
                }

                Excel.Range rngData = _ws.Range[rng.ToString()];
                
                if (info["ValoreDefault"] != DBNull.Value) 
                {
                    rngData.Value = info["ValoreDefault"];
                }
                else if (info["FormulaInCella"].Equals("1"))
                {
                    int deltaNeg;
                    int deltaPos;
                    string formula = "=" + PreparaFormula(info, "DATA0", "DATA1", 24, out deltaNeg, out deltaPos);

                    if (info["SiglaTipologiaInformazione"].Equals("OTTIMO"))
                    {
                        rngData.Cells[1].Formula = "=SUM(" + rng.Columns[1, rng.Columns.Count - 1] + ")";
                        deltaNeg = 1;
                    }
                    _ws.Range[rng.Columns[deltaNeg, rng.Columns.Count - 1 - deltaPos].ToString()].Formula = formula;
                }

                if (info["ValoreData0H24"] != DBNull.Value)
                    rngData.Cells[1].Value = info["ValoreData0H24"];
            }
        }
        /// <summary>
        /// Inserisce i parametri.
        /// </summary>
        protected virtual void InsertParametri()
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
            DataView parametriD = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PARAMETRO_D].DefaultView;
            DataView parametriH = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PARAMETRO_H].DefaultView;

            foreach (DataRowView info in informazioni)
            {
                object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                CicloGiorni((oreGiorno, suffissoData, giorno) =>
                {
                    Range rngData = _definedNames.Get(siglaEntita, info["SiglaInformazione"], suffissoData).Extend(colOffset: oreGiorno);
                    parametriD.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaParametro = '" + info["SiglaTipologiaParametro"] + "' AND DataIV <= '" + giorno.ToString("yyyyMMdd") + "' AND DataFV >= '" + giorno.ToString("yyyyMMdd") + "'";

                    if (parametriD.Count > 0)
                        _ws.Range[rngData.ToString()].Formula = parametriD[0]["Valore"];
                    else
                    {
                        parametriH.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaParametro = '" + info["SiglaTipologiaParametro"] + "' AND DataIV <= '" + giorno.ToString("yyyyMMdd") + "' AND DataFV >= '" + giorno.ToString("yyyyMMdd") + "'";

                        object[] values = parametriH.ToTable(false, "Valore").AsEnumerable().Select(r => r["Valore"]).ToArray();

                        if (values.Length > 0)
                            _ws.Range[rngData.ToString()].Value = values;
                    }
                });
            }
        }
        /// <summary>
        /// Crea la formattazione condizionale.
        /// </summary>
        protected virtual void FormattazioneCondizionale()
        {
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
            DataView formattazione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE_FORMATTAZIONE].DefaultView;
            int colOffset = _definedNames.GetColOffset(_dataFine);
            foreach (DataRowView info in informazioni)
            {
                object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                
                int offsetAdjust = (_struttura.visData0H24 && info["Data0H24"].Equals("0") ? 1 : 0);
                Range rng = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, info["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), _definedNames.GetFirstCol() + offsetAdjust, 1, colOffset - offsetAdjust);

                Excel.Range rngData = _ws.Range[rng.ToString()];

                formattazione.RowFilter = (info["SiglaEntitaRif"] is DBNull ? "SiglaEntita" : "SiglaEntitaRif") + " = '" + siglaEntita + "' AND SiglaInformazione = '" + info["SiglaInformazione"] + "'";
                foreach (DataRowView format in formattazione)
                {
                    
                    string[] valore = format["Valore"].ToString().Replace("\"", "").Split('|');
                    if (format["NomeCella"] != DBNull.Value)
                    {
                        int refRow = _definedNames.GetRowByNameSuffissoData(siglaEntita, format["NomeCella"], Date.GetSuffissoData(_dataInizio));
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
        /// <summary>
        /// Inserisce i grafici creando anche le serie.
        /// </summary>
        protected virtual void InsertGrafici()
        {
            DataView grafici = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_GRAFICO].DefaultView;
            DataView graficiInfo = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_GRAFICO_INFORMAZIONE].DefaultView;

            int i = 1;
            int col = _definedNames.GetColData1H1();
            int colOffset = _definedNames.GetColOffset(_dataFine) - (_struttura.visData0H24 ? 1 : 0);
            foreach (DataRowView grafico in grafici)
            {
                SplashScreen.UpdateStatus("Genero grafici");
                string name = DefinedNames.GetName(grafico["SiglaEntita"], "GRAFICO" + i++, Struct.tipoVisualizzazione == "O" ? "" : Date.GetSuffissoData(_dataInizio));

                Range rngGrafico = new Range(_definedNames.GetRowByName(name), col, 1, colOffset);
                Excel.Range xlRngGrafico = _ws.Range[rngGrafico.ToString()];
                xlRngGrafico.Merge();
                xlRngGrafico.Style = "Area grafici";
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
                chart.PlotVisibleOnly = false;

                chart.PlotArea.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                string rowFilter = graficiInfo.RowFilter;
                graficiInfo.RowFilter = rowFilter + " AND SiglaGrafico = '" + grafico["SiglaGrafico"] + "'";

                foreach (DataRowView info in graficiInfo)
                {
                    Range rngDati = new Range(_definedNames.GetRowByNameSuffissoData(grafico["SiglaEntita"], info["SiglaInformazione"], Date.GetSuffissoData(_dataInizio)), col, 1, colOffset);
                    Excel.Series serie = chart.SeriesCollection().NewSeries();
                    serie.Name = info["DesInformazione"].ToString();
                    serie.Values = _ws.Range[rngDati.ToString()];
                    serie.ChartType = (Excel.XlChartType)info["ChartType"];
                    serie.Interior.ColorIndex = info["InteriorColor"];
                    serie.Border.ColorIndex = info["BorderColor"];
                    serie.Border.Weight = info["BorderWeight"];
                    serie.Border.LineStyle = info["BorderLineStyle"];
                }
                graficiInfo.RowFilter = rowFilter;
            }
        }
        /// <summary>
        /// Aggiorna tutti i grafici del foglio.
        /// </summary>
        public override void AggiornaGrafici()
        {
            if (_ws.ChartObjects().Count > 0)
            {
                _ws.Calculate();
                Excel.ChartObjects charts = _ws.ChartObjects();
                foreach (Excel.ChartObject chart in charts)
                {
                    int col;
                    if (chart.Name.Contains("DATA"))
                    {
                        col = _definedNames.GetColFromDate(chart.Name.Split(Simboli.UNION[0]).Last());
                    }
                    else
                    {
                        col = _definedNames.GetColFromDate();
                    }
                    int row = _definedNames.GetRowByName(chart.Name);
                    Excel.Range rng = _ws.Range[Range.GetRange(row, col)];
                    AggiornaGrafici(chart.Chart, rng.MergeArea);
                    chart.Chart.Refresh();
                }
            }
        }
        /// <summary>
        /// Allinea il grafico al range in modo da far combaciare la barra delle ordinate con la prima colonna dell'area dati. Per far questo calcola la dimensione in punti dei label di ordinata e sposta di conseguenza l'area del grafico.
        /// </summary>
        /// <param name="chart">Microsoft.Office.Interop.Excel.Chart da aggiornare.</param>
        /// <param name="rigaGrafico">Microsoft.Office.Interop.Excel.Range a cui il grafico appartiene.</param>
        private void AggiornaGrafici(Excel.Chart chart, Excel.Range rigaGrafico)
        {
            SplashScreen.UpdateStatus("Aggiorno grafici " + chart.Name);

            //calcolo i valori max e min per aggiornare la scala
            bool allNull = true;
            double minValue = double.MaxValue;
            double maxValue = double.MinValue;

            foreach (Excel.Series s in chart.SeriesCollection())
            {
                Array val = s.Values as Array;

                if (val.OfType<double>().Any())
                {
                    allNull = false;
                    minValue = Math.Min(minValue, val.Cast<double>().Min());
                    maxValue = Math.Max(maxValue, val.Cast<double>().Max());
                }
            }

            if (!allNull)
            {
                chart.Axes(Excel.XlAxisType.xlValue).MaximumScale = Math.Round(maxValue + maxValue * 5 / 100);
                chart.Axes(Excel.XlAxisType.xlValue).MinimumScale = Math.Round(minValue - minValue * 5 / 100);
            }

            //resize dell'area del grafico per adattarla alle ore
            Graphics grfx = Graphics.FromImage(new Bitmap(1, 1));
            grfx.PageUnit = GraphicsUnit.Point;
            float sizeMax = float.MinValue;

            for(double val = chart.Axes(Excel.XlAxisType.xlValue).MinimumScale; val <= chart.Axes(Excel.XlAxisType.xlValue).MaximumScale; val += chart.Axes(Excel.XlAxisType.xlValue).MajorUnit) 
            {
                SizeF tmpSize = grfx.MeasureString(val.ToString(), new Font("Verdana", 11));
                sizeMax = Math.Max(sizeMax, tmpSize.Width);
            }

            //MANTENERE ORDINE DI QUESTE ISTRUZIONI
            chart.ChartArea.Left = rigaGrafico.Left - sizeMax - 7;      //sposto a destra il grafico
            chart.ChartArea.Width = rigaGrafico.Width + sizeMax + 4;    //aumento la larghezza del grafico
            chart.PlotArea.InsideLeft = 0d;                             //allineo il grafico al bordo sinistro dell'area esterna al grafico
            chart.PlotArea.Width = chart.ChartArea.Width + 3;           //aumento la larghezza dell'area esterna al grafico
        }

        #endregion

        /// <summary>
        /// Launcher per caricare le informazioni e i commenti dal DB.
        /// </summary>
        public override void CaricaInformazioni()
        {
            try
            {
                if (DataBase.OpenConnection())
                {
                    DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
                    categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "'";

                    _dataInizio = DataBase.DB.DataAttiva;

                    DateTime dataFineMax = _dataInizio.AddDays(_intervalloGiorniMax);

                    DataView datiApplicazioneH = DataBase.LocalDB.Tables[DataBase.Tab.DATI_APPLICAZIONE_H].DefaultView;
                    DataView insertManuali = DataBase.LocalDB.Tables[DataBase.Tab.DATI_APPLICAZIONE_COMMENTO].DefaultView;

                    if (Struct.tipoVisualizzazione == "O")
                    {
                        foreach (DataRowView entita in categoriaEntita)
                        {
                            object siglaEntita = entita["Gerarchia"] is DBNull ? entita["SiglaEntita"] : entita["Gerarchia"];
                            SplashScreen.UpdateStatus("Scrivo informazioni " + entita["DesEntita"]);
                            datiApplicazioneH.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND CONVERT(Data, System.Int32) <= " + _dataFineUP[siglaEntita].ToString("yyyyMMdd");
                            CaricaInformazioniEntita(datiApplicazioneH);
                            insertManuali.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND CONVERT(SUBSTRING(Data, 1, 8), System.Int32) <= " + _dataFineUP[siglaEntita].ToString("yyyyMMdd");
                            CaricaCommentiEntita(insertManuali);
                        }
                    }
                    else
                    {
                        datiApplicazioneH.RowFilter = "SiglaCategoria='" + _siglaCategoria + "'";
                        CaricaInformazioniEntita(datiApplicazioneH);
                        CaricaCommentiEntita(insertManuali);
                    }

                    //carico dati giornalieri
                    DataView datiApplicazioneD = DataBase.LocalDB.Tables[DataBase.Tab.DATI_APPLICAZIONE_D].DefaultView;

                    foreach (DataRowView dato in datiApplicazioneD)
                    {
                        if (_definedNames.IsDefined(dato["SiglaEntita"]))
                        {
                            Range rng = new Range(_definedNames.GetRowByNameSuffissoData(dato["SiglaEntita"], dato["SiglaInformazione"], Date.GetSuffissoData(dato["Data"].ToString())), _definedNames.GetFirstCol() - 1);

                            _ws.Range[rng.ToString()].Value = dato["Valore"];
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "CaricaInformazioni [all = 1]: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// Carica le informazioni.
        /// </summary>
        /// <param name="datiApplicazione">Tabella contenente tutte le informazioni da scrivere.</param>
        protected virtual void CaricaInformazioniEntita(DataView datiApplicazione)
        {
            //SplashScreen.UpdateStatus("Scrivo informazioni");
            foreach (DataRowView dato in datiApplicazione)
            {
                DateTime giorno = DateTime.ParseExact(dato["Data"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                
                if (giorno <= _dataFineUP[dato["SiglaEntita"]])
                {
                    
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
                        _ws.Range[rng.ToString()].Value = o.ToArray();

                        if (giorno == DataBase.DataAttiva && Regex.IsMatch(dato["SiglaInformazione"].ToString(), @"RIF\d+"))
                        {
                            Selection s = _definedNames.GetSelectionByRif(rng);
                            s.ClearSelections(_ws);
                            s.Select(_ws, int.Parse(o[0].ToString().Split('.')[0]));
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Carica i commenti.
        /// </summary>
        /// <param name="insertManuali">Tabella che contiene tutte le informazioni che necessitano del commento</param>
        protected virtual void CaricaCommentiEntita(DataView insertManuali)
        {
            foreach (DataRowView commento in insertManuali)
            {
                DateTime giorno = DateTime.ParseExact(commento["Data"].ToString().Substring(0,8), "yyyyMMdd", CultureInfo.InvariantCulture);

                if (giorno <= _dataFineUP[commento["SiglaEntita"]])
                {
                    SplashScreen.UpdateStatus("Scrivo commenti " + commento["SiglaEntita"]);
                    Range rng = _definedNames.Get(commento["SiglaEntita"], commento["SiglaInformazione"], Date.GetSuffissoData(giorno), Date.GetSuffissoOra(commento["Data"].ToString()));
                    _ws.Range[rng.ToString()].ClearComments();
                    _ws.Range[rng.ToString()].AddComment("Valore inserito manualmente");
                }
            }
        }        

        /// <summary>
        /// Prepara la formula per essere scritta in cella. Sostituisce i parametri con i riferimenti delle celle corrispondenti. Calcola anche un possibile offset se la formula va a controllare valori delle ore precedenti o successive.
        /// </summary>
        /// <param name="info">La riga dell'informazione.</param>
        /// <param name="suffissoDataPrec">Il suffisso della data antecedente a quella in cui si sta lavorando (solitamente DATA0).</param>
        /// <param name="suffissoData">Il suffisso della data in cui si sta lavorando (solitamente DATA1).</param>
        /// <param name="oreDataPrec">Numero di ore della data precedente.</param>
        /// <param name="deltaNeg">Parametro di output che indica l'offset dall'inizio del giorno.</param>
        /// <param name="deltaPos">Parametro di output che indica l'offset dalla fine del giorno.</param>
        /// <returns></returns>
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
                        DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
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
                    Range rng = _definedNames.Get(siglaEntita, siglaInformazione, suffData, suffOra);

                    return rng.ToString();
                }, RegexOptions.IgnoreCase);
                return formula;
            }
            deltaNeg = 0;
            deltaPos = 0;

            return "";
        }
        /// <summary>
        /// Launcher per l'aggiornamento dei dati.
        /// </summary>
        public override void UpdateData()
        {
            SplashScreen.UpdateStatus("Aggiorno informazioni");
            CancellaDati();
            AggiornaDateTitoli();
            CaricaParametri();
            CaricaInformazioni();
            AggiornaGrafici();
        }
        #region UpdateData

        /// <summary>
        /// Cancella le informazioni da aggiornare in tutti i giorni.
        /// </summary>
        private void CancellaDati()
        {
            CancellaDati(DataBase.DataAttiva, true);
        }
        /// <summary>
        /// Cancella le informazioni a partire da giorno. Se all è a true, cancella tutti giorni successivi altrimenti cancella il solo giorno.
        /// </summary>
        /// <param name="giorno">Data di partenza della cancellazione.</param>
        /// <param name="all">Se true cancella tutti i dati a partire dalla data di partenza.</param>
        private void CancellaDati(DateTime giorno, bool all = false)
        {
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "'"; // AND (Gerarchia = '' OR Gerarchia IS NULL )";

            string suffissoData = Date.GetSuffissoData(giorno);
            int colOffset = _definedNames.GetColOffset();
            if (!all)
                colOffset = Date.GetOreGiorno(giorno);

            foreach (DataRowView entita in categoriaEntita)
            {
                SplashScreen.UpdateStatus("Cancello dati " + entita["DesEntita"]);
                DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND FormulaInCella = '0'";// AND ValoreDefault IS NULL";

                foreach (DataRowView info in informazioni)
                {
                    int col = all ? _definedNames.GetFirstCol() : _definedNames.GetColFromDate(suffissoData);
                    int realColOffset = colOffset;
                    object siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                    if (Struct.tipoVisualizzazione == "O")
                    {
                        int row = _definedNames.GetRowByName(siglaEntita, info["SiglaInformazione"]);
                        if (info["SiglaTipologiaInformazione"].Equals("GIORNALIERA"))
                        {
                            Excel.Range rngData = _ws.Range[Range.GetRange(row, col - 1)];
                            rngData.Value = "";
                        }
                        else
                        {
                            if (_struttura.visData0H24 && info["Data0H24"].Equals("0"))
                            {
                                col++;
                                realColOffset--;
                            }
                            Excel.Range rngData = _ws.Range[Range.GetRange(row, col, 1, realColOffset)];
                            //rngData.Value = info["ValoreDefault"] != DBNull.Value ? info["ValoreDefault"] : "";
                            rngData.Value = "";

                            rngData.ClearComments();
                            Style.RangeStyle(rngData, backColor: info["BackColor"], foreColor: info["ForeColor"]);
                        }                        
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
                            SplashScreen.UpdateStatus("Cancello dati " + g.ToShortDateString());

                            int row = _definedNames.GetRowByNameSuffissoData(siglaEntita, info["SiglaInformazione"], suffData);
                            if (info["SiglaTipologiaInformazione"].Equals("GIORNALIERA"))
                            {
                                Excel.Range rngData = _ws.Range[Range.GetRange(row, col - 1)];
                                rngData.Value = "";
                            }
                            else
                            {
                                Excel.Range rng = _ws.Range[Range.GetRange(row, col, 1, oreGiorno)];
                                rng.Value = "";
                                rng.ClearComments();
                                Style.RangeStyle(rng, backColor: info["BackColor"], foreColor: info["ForeColor"]);
                            }
                        });
                    }
                }
                //reset colonna 24esima 25esima ora
                if (all && Struct.tipoVisualizzazione == "V" && informazioni.Count > 0)
                {
                    DateTime dataInizio = DataBase.DataAttiva;
                    DateTime dataFine = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni);

                    object siglaEntita = informazioni[0]["SiglaEntitaRif"] is DBNull ? informazioni[0]["SiglaEntita"] : informazioni[0]["SiglaEntitaRif"];

                    CicloGiorni(dataInizio, dataFine, (oreGiorno, suffData, g) =>
                    {
                        Range rngData = new Range(_definedNames.GetRowByNameSuffissoData(siglaEntita, informazioni[0]["SiglaInformazione"], suffData), _definedNames.GetFirstCol(), informazioni.Count, oreGiorno);                        

                        int ore = Date.GetOreGiorno(g);
                        if (ore == 23) 
                        {
                            _ws.Range[rngData.Columns[rngData.Columns.Count - 2, rngData.Columns.Count - 1].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                        }
                        else if (ore == 24)
                        {
                            _ws.Range[rngData.Columns[rngData.Columns.Count - 2].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternNone;
                            _ws.Range[rngData.Columns[rngData.Columns.Count - 1].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                        }
                        else if (ore == 25)
                        {
                            _ws.Range[rngData.Columns[rngData.Columns.Count - 2].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternNone;
                            _ws.Range[rngData.Columns[rngData.Columns.Count - 1].ToString()].Interior.Pattern = Excel.XlPattern.xlPatternNone;
                        }
                    });
                }
            }
        }
        /// <summary>
        /// Aggiorna le date dei titolo (per il caso in cui l'aggiornamento venga da un cambio giorno).
        /// </summary>
        public override void AggiornaDateTitoli()
        {
            if (Struct.tipoVisualizzazione == "O")
            {
                int row = _struttura.rigaBlock - 2;
                for (int i = 0; i < _definedNames.DaySuffx.Length; i++)
                {
                    if (_definedNames.DaySuffx[i] != "DATA0")
                    {
                        int col = _definedNames.GetColFromDate(_definedNames.DaySuffx[i]);
                        _ws.Range[Range.GetRange(row, col)].Value = Date.GetDataFromSuffisso(_definedNames.DaySuffx[i]);
                    }
                }
            }
            else
            {
                DefinedNames gotos = new DefinedNames(_ws.Name, DefinedNames.InitType.GOTOsThisSheet);

                for (int i = 0; i <= Struct.intervalloGiorni; i++)
                {
                    DateTime giorno = DataBase.DataAttiva.AddDays(i);
                    string suffissoData = Date.GetSuffissoData(giorno);
                    
                    int row = _definedNames.GetRowByName(suffissoData, "T");
                    int col = _definedNames.GetFirstCol();
                    _ws.Range[Range.GetRange(row, col)].Value = giorno;

                    row += 2;//_definedNames.GetRowByName(suffissoData, "TITOLO_VERTICALE");
                    col -= (_visParametro + 1);
                    if (_ws.Range[Range.GetRange(row, col)].Value != null)
                        _ws.Range[Range.GetRange(row, col)].Value = giorno;

                    _ws.Range[gotos.GetFromAddressGOTO(i)].Value = giorno;

                }
            }
        }
        /// <summary>
        /// Carica i parametri e valori di default.
        /// </summary>
        protected void CaricaParametri()
        {
            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            DataView informazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;

            categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "' AND (Gerarchia = '' OR Gerarchia IS NULL )";
            _dataInizio = DataBase.DB.DataAttiva;

            foreach (DataRowView entita in categoriaEntita)
            {
                SplashScreen.UpdateStatus("Carico parametri " + entita["DesEntita"]);
                _dataFine = _dataFineUP[entita["SiglaEntita"]];

                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaTipologiaParametro IS NOT NULL";
                InsertParametri();

                SplashScreen.UpdateStatus("Aggiorno valori di default " + entita["DesEntita"]);
                informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND ValoreDefault IS NOT NULL";
                InsertFormuleValoriDefault();
            }
        }

        #endregion

        public void Dispose()
        {
            if (!_disposed)
            {
                GC.SuppressFinalize(this);
                _disposed = true;
            }
        }

        #endregion
    }
}
