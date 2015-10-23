using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Core;
using Iren.ToolsExcel.UserConfig;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Iren.ToolsExcel.Utility
{
    public class Workbook
    {
        #region Variabili

        /// <summary>
        /// Il workbook.
        /// </summary>
        protected static IToolsExcelThisWorkbook _wb;
        /// <summary>
        /// Flag che viene utilizzato per bloccare l'evento SheetSelectionChange quando la selezione è cambiata dal pannello laterale dei check.
        /// </summary>
        public static bool FromErrorPane = false;

        public static IWin32Window Window;

        #endregion

        #region Proprietà

        /// <summary>
        /// L'oggetto Excel del Workbook per accedere a tutti gli handler e proprietà. (Read only)
        /// </summary>
        public static Microsoft.Office.Tools.Excel.Workbook WB { get { return _wb.Base; } }
        /// <summary>
        /// Scorciatoia per accedere all'oggetto Excel del foglio Main (sempre presente in tutti i fogli).
        /// </summary>
        public static Microsoft.Office.Tools.Excel.Worksheet Main { get { return _wb.Main; } }
        /// <summary>
        /// Scorciatoia per accedere all'oggetto Excel del foglio di Log (sempre presente in tutti i fogli).
        /// </summary>
        public static Microsoft.Office.Tools.Excel.Worksheet Log { get { return _wb.Log; } }
        /// <summary>
        /// Scorciatoia per accedere al foglio attivo.
        /// </summary>
        public static Microsoft.Office.Tools.Excel.Worksheet ActiveSheet { get { return _wb.ActiveSheet; } }
        /// <summary>
        /// Scorciatoia per accedere all'oggetto Application di Excel.
        /// </summary>
        public static Excel.Application Application { get { return _wb.Application; } }
        /// <summary>
        /// Lista di tutti i fogli che rappresentano una Categoria sul DB (non fanno parte i fogli Log, Main, MSDx). I fogli non sono indicizzati per nome, solo per indice.
        /// </summary>
        public static IList<Excel.Worksheet> CategorySheets { get { return _wb.Sheets.Cast<Excel.Worksheet>().Where(ws => ws.Name != "Log" && ws.Name != "Main" && !ws.Name.StartsWith("MSD")).ToList(); } }
        /// <summary>
        /// Lista di tutti fogli indicizzati per nome.
        /// </summary>
        public static Excel.Sheets Sheets { get { return WB.Sheets; } }
        /// <summary>
        /// Lista dei folgi MSDx utile solo in Invio Programmi.
        /// </summary>
        public static IList<Excel.Worksheet> MSDSheets { get { return _wb.Sheets.Cast<Excel.Worksheet>().Where(ws => ws.Name.StartsWith("MSD")).ToList(); } }
        /// <summary>
        /// La versione dell'applicazione.
        /// </summary>
        public static System.Version WorkbookVersion { get { return _wb.Version; } }
        /// <summary>
        /// La versione della classe Core
        /// </summary>
        public static System.Version CoreVersion { get { return DataBase.DB.GetCurrentV(); } }
        /// <summary>
        /// La versione della classe Base.
        /// </summary>
        public static System.Version BaseVersion { get { return Assembly.GetExecutingAssembly().GetName().Version; } }
        /// <summary>
        /// Flag per attivare/disattivare il refresh dello schermo.
        /// </summary>
        public static bool ScreenUpdating { get { return Application.ScreenUpdating; } set { Application.ScreenUpdating = value; } }



        public static Utility.Repository Repository { get; private set; }


        //public static DateTime DataAttiva { get { return _wb.DataAttiva; } }

        #endregion

        #region Metodi

        /// <summary>
        /// Carica dal DB i dati riguardanti le proprietà dell'applicazione che si trovano nella tabella APPLICAZIONE. Assegna alle variabili globali di applicazione i valori.
        /// </summary>
        public static void AggiornaParametriApplicazione()
        {
            DataRow r = Workbook.Repository.CaricaApplicazione(_wb.IdApplicazione);
            if (r == null)
                throw new ApplicationNotFoundException("L'appID inserito non ha restituito risultati.");

            Simboli.nomeApplicazione = r["DesApplicazione"].ToString();
            Struct.intervalloGiorni = (r["IntervalloGiorniEntita"] is DBNull ? 0 : (int)r["IntervalloGiorniEntita"]);
            Struct.tipoVisualizzazione = r["TipoVisualizzazione"] is DBNull ? "O" : r["TipoVisualizzazione"].ToString();
            Struct.visualizzaRiepilogo = r["VisRiepilogo"] is DBNull ? true : r["VisRiepilogo"].Equals("1");

            Struct.cell.width.empty = double.Parse(r["ColVuotaWidth"].ToString());
            Struct.cell.width.dato = double.Parse(r["ColDatoWidth"].ToString());
            Struct.cell.width.entita = double.Parse(r["ColEntitaWidth"].ToString());
            Struct.cell.width.informazione = double.Parse(r["ColInformazioneWidth"].ToString());
            Struct.cell.width.unitaMisura = double.Parse(r["ColUMWidth"].ToString());
            Struct.cell.width.parametro = double.Parse(r["ColParametroWidth"].ToString());
            Struct.cell.width.jolly1 = double.Parse(r["ColJolly1Width"].ToString());
            Struct.cell.height.normal = double.Parse(r["RowHeight"].ToString());
            Struct.cell.height.empty = double.Parse(r["RowVuotaHeight"].ToString());
        }
        /// <summary>
        /// Imposta il mercato attivo in base all'orario. Se necessario cambia anche la data attiva e imposta il foglio come da aggiornare.
        /// </summary>
        /// <param name="appID">L'ID applicazione che identifica anche in quale mercato il foglio è impostato.</param>
        /// <param name="dataAttiva">La data attiva da modificare all'occorrenza.</param>
        /// <returns>Restituisce true se il foglio è da aggiornare, false altrimenti.</returns>
        private static bool SetMercato()
        {
            int idApplicazioneOLD = _wb.IdApplicazione;
            DateTime dataAttivaOld = _wb.DataAttiva;

            //configuro la data attiva
            int ora = DateTime.Now.Hour;
            if (ora > 17)
                _wb.DataAttiva = DateTime.Today.AddDays(1);
            else if (ora >= 7 && ora <= 17)
                _wb.DataAttiva = DateTime.Today;

            //configuro il mercato attivo
            string[] mercatiDisp = Workbook.AppSettings("Mercati").Split('|');
            string[] appIDs = Workbook.AppSettings("AppIDMSD").Split('|');
            for (int i = 0; i < mercatiDisp.Length; i++)
            {
                string[] ore = Workbook.AppSettings("Ore" + mercatiDisp[i]).Split('|');
                if (ore.Contains(ora.ToString()))
                {
                    _wb.IdApplicazione = int.Parse(appIDs[i]);
                    break;
                }
            }

            Simboli.AppID = _wb.IdApplicazione.ToString();

            if (_wb.IdApplicazione != idApplicazioneOLD || dataAttivaOld != _wb.DataAttiva)
            {
                Workbook.ChangeAppSettings("DataAttiva", _wb.DataAttiva.ToString("yyyyMMdd"));
                Simboli.AppID = _wb.DataAttiva.ToString();

                return true;
            }

            return false;
        }
        /// <summary>
        /// Aggiorna la data per le applicazione Validazione TL e Previsione CT.
        /// </summary>
        /// <param name="appID">L'ID applicazione</param>
        /// <param name="dataAttiva">La data attiva da cambiare se necessario</param>
        /// <returns>Restituisce true se il foglio è da aggiornare, false altrimenti.</returns>
        private static bool AggiornaData()
        {
            DateTime dataAttivaOld = _wb.DataAttiva;

            if (_wb.IdApplicazione == 12)
            {
                //configuro la data attiva
                int ora = DateTime.Now.Hour;
                if (ora <= 15)
                    _wb.DataAttiva = DateTime.Today.AddDays(1);
                else
                    _wb.DataAttiva = DateTime.Today.AddDays(2);
            }
            else
            {
                _wb.DataAttiva = DateTime.Today.AddDays(-1);
            }


            if (dataAttivaOld != _wb.DataAttiva)
            {
                Workbook.ChangeAppSettings("DataAttiva", _wb.DataAttiva.ToString("yyyyMMdd"));
                return true;
            }
            return false;
        }
        /// <summary>
        /// Aggiorna i label indicanti lo stato dei Database in seguito ad un cambio di stato.
        /// </summary>
        public static void AggiornaLabelStatoDB()
        {
            //disabilito l'aggiornamento in caso di modifica dati... lo ripeto alla chiusura in caso
            if (!Simboli.ModificaDati)
            {
                bool isProtected = true;
                try
                {
                    Workbook.WB.Application.ScreenUpdating = false;
                    isProtected = Main.ProtectContents;

                    if (isProtected)
                        Main.Unprotect(Simboli.pwd);


                    Riepilogo main = new Riepilogo(Utility.Workbook.Main);

                    if (DataBase.OpenConnection())
                    {
                        Dictionary<Core.DataBase.NomiDB, ConnectionState> stato = DataBase.StatoDB;
                        Simboli.SQLServerOnline = stato[Core.DataBase.NomiDB.SQLSERVER] == ConnectionState.Open;
                        Simboli.ImpiantiOnline = stato[Core.DataBase.NomiDB.IMP] == ConnectionState.Open;
                        Simboli.ElsagOnline = stato[Core.DataBase.NomiDB.ELSAG] == ConnectionState.Open;

                        main.UpdateData();

                        DataBase.CloseConnection();
                    }
                    else
                    {
                        Simboli.SQLServerOnline = false;
                        Simboli.ImpiantiOnline = false;
                        Simboli.ElsagOnline = false;

                        main.RiepilogoInEmergenza();
                    }

                    if (isProtected)
                        Main.Protect(Simboli.pwd);
                }
                catch { }

                //lo faccio a parte perché se andasse in errore prima deve almeno provare a riattivare lo screen updating!!!
                try { Workbook.WB.Application.ScreenUpdating = true; }
                catch { }
            }
        }
        /// <summary>
        /// Handler per il PropertyChanged della classe Core.DataBase. Attiva l'aggiornamento dei label.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void _db_StatoDBChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            AggiornaLabelStatoDB();
        }
        /// <summary>
        /// Scrive il DataSet allo stato attuale nelle custom parts del foglio. In questo modo la persistenza è garantita. Le tabelle di Log e Modifica saranno sempre vuote a questo punto.
        /// </summary>
        public static void DumpDataSet()
        {
            //creo gli stream necessari a scrivere l'XML in modo compatto (senza formattazione) per risparmiare spazio e velocizzare il processo.
            StringWriter strWriter = new StringWriter();
            XmlWriter xmlWriter = XmlWriter.Create(strWriter);
            //scrivo il dataset LocalDB con annesso lo schema (è successo che senza lo schema, quando tutta una colonna è vuota, non viene scritta nell'XML e questo crea problemi quando poi si tenta di caricare nuovi dati).
            Utility.DataBase.LocalDB.WriteXml(xmlWriter, XmlWriteMode.WriteSchema);

            XElement root = XElement.Parse(strWriter.ToString());
            XNamespace ns = WB.Name;

            //cancello tutte le righe della tabella di log che verrà riempita ad ogni avvio/modifica del log.
            IEnumerable<XElement> log =
                from tables in root.Elements(ns + Utility.DataBase.Tab.LOG)
                select tables;

            log.Remove();

            string locDBXml = strWriter.ToString();
            Microsoft.Office.Core.CustomXMLPart part;

            //Non avendo trovato un modo di interagire con le custom parts, provo a cancellare quella esistente (se esiste, altrimenti va in errore e aggiunge la nuova senza problemi).
            //try { _wb.CustomXMLParts[WB.Name].Delete(); }
            //catch { }
            //part = _wb.CustomXMLParts.Add();
            //carico nella nuova custom part il contenuto.
            //part.LoadXML(root.ToString(SaveOptions.DisableFormatting));
        }
        /// <summary>
        /// Restituisce lo UserConfigElement collegato alla chiave configKey nella sezione usrConfig (da non confondere con appSettings).
        /// </summary>
        /// <param name="configKey">Chiave.</param>
        /// <returns>Restituisce l'elemento ricercato.</returns>
        public static UserConfiguration GetUsrConfiguration()
        {
            return (UserConfiguration)ConfigurationManager.GetSection("usrConfig");
        }
        public static UserConfigElement GetUsrConfigElement(string configKey)
        {
            var settings = GetUsrConfiguration();
            return (UserConfigElement)settings.Items[configKey];
        }

        /// <summary>
        /// Restituisce un array con le tre componenti intere Red Green Blue a partire da una stringa suddivisa con un separatore sep. Non ha una gestione di errore, se il parser non riesce ad interpretare la stringa, va in errore.
        /// </summary>
        /// <param name="rgb">Stringa nel formato RRR[sep]GGG[sep]BBB.</param>
        /// <param name="sep">Separatore.</param>
        /// <returns>Restituisce le tre componenti trovate.</returns>
        public static int[] GetRGBFromString(string rgb, char sep = ';')
        {
            string[] rgbComp = rgb.Split(sep);

            return new int[] { int.Parse(rgbComp[0]), int.Parse(rgbComp[1]), int.Parse(rgbComp[2]) };
        }

        #region AppSettings
        /// <summary>
        /// Quando il file è criptato capita che senza il refresh vada in errore.
        /// </summary>
        /// <param name="key">La chiave da ricercare nella sezione appSettings</param>
        /// <returns>Restituisce la stringa del Value.</returns>
        public static string AppSettings(string key)
        {
            try
            {
                return ConfigurationManager.AppSettings[key];
            }
            catch
            {
                ConfigurationManager.RefreshSection("appSettings");
                return ConfigurationManager.AppSettings[key];
            }
        }
        /// <summary>
        /// Assegna value al valore della chiave key della sesione appSettings del file di configurazione. Alla fine dell'operazione esegue il refresh della sezione in modo da forzare la riscrittura su disco dei nuovi valori.
        /// </summary>
        /// <param name="key">Chiave da modificare.</param>
        /// <param name="value">Nuovo valore da assegnare.</param>
        public static void ChangeAppSettings(string key, string value)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings[key].Value = value;
            config.Save(ConfigurationSaveMode.Minimal);
            ConfigurationManager.RefreshSection("appSettings");
        }
        #endregion

        #region Init
        public static void InitLog()
        {
            DataTable dtLog = DataBase.Select(DataBase.SP.APPLICAZIONE_LOG);
            if (dtLog != null)
            {
                dtLog.TableName = DataBase.Tab.LOG;
                if (DataBase.LocalDB.Tables.Contains(DataBase.Tab.LOG))
                    DataBase.LocalDB.Tables[DataBase.Tab.LOG].Merge(dtLog);
                else
                    DataBase.LocalDB.Tables.Add(dtLog);

                DataView dv = DataBase.LocalDB.Tables[DataBase.Tab.LOG].DefaultView;
                dv.Sort = "Data DESC";
            }
        }
        private static void InitUtente()
        {
            DataTable dtUtente = DataBase.Select(DataBase.SP.UTENTE, new QryParams() { { "@CodUtenteWindows", Environment.UserName } });
            if (dtUtente != null)
            {
                if(dtUtente.Rows.Count == 0)
                {
                    _wb.IdUtente = 0;
                    _wb.NomeUtente = "NON CONFIGURATO";
                }
                else
                {
                    _wb.IdUtente = (int)dtUtente.Rows[0]["IdUtente"];
                    _wb.NomeUtente = dtUtente.Rows[0]["Nome"].ToString();
                }

                return;
            }

            System.Windows.Forms.MessageBox.Show("Errore durante l'inizializzazione dell'utente.", Simboli.nomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
        }
        private static bool Initialize()
        {
            //CryptHelper.CryptSection("connectionStrings", "appSettings");
            //controllo le aree di rete (se presenti)
            var usrConfig = GetUsrConfiguration();
            Dictionary<string, string> pathNonDisponibili = new Dictionary<string, string>();
            foreach (UserConfigElement ele in usrConfig.Items)
            {
                if (ele.Type == UserConfigElement.ElementType.path)
                {
                    string pathStr = ele.Value;

                    try { System.Security.AccessControl.DirectorySecurity ds = Directory.GetAccessControl(pathStr); }
                    catch { pathNonDisponibili.Add(ele.Desc, pathStr); }
                }
            }
            //segnalo all'utente l'impossibilità di accedere alle aree di rete
            if (pathNonDisponibili.Count > 0)
            {
                string paths = "\n";
                foreach (var kv in pathNonDisponibili)
                    paths += " - " + kv.Key + " : '" + kv.Value + "'\n";

                System.Windows.Forms.MessageBox.Show("I path seguenti non sono raggiungibili o non presentano privilegi di scrittura:" + paths, Simboli.nomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }


            DataBase.Initialize(_wb.Ambiente);
            DataBase.DB.PropertyChanged += _db_StatoDBChanged;
            //DataBase.InitNewLocalDB();

            //bool localDBNotPresent = Repository.Tables.Count == 0;


            //try
            //{
            //    Office.CustomXMLPart xmlPart = WB.CustomXMLParts[WB.Name];
            //    StringReader sr = new StringReader(xmlPart.XML);
            //    DataBase.LocalDB.ReadXml(sr);
            //}
            //catch
            //{
            //    localDBNotPresent = true;
            //    DataBase.LocalDB.Namespace = WB.Name;
            //}

            bool toUpdate = Repository.TablesCount == 0;

            //per Invio Programmi
            if (AppSettings("Mercati") != null)
                toUpdate = SetMercato() || toUpdate;

            //per Previsione Carico Termico & Validazione Teleriscaldamento
            if (_wb.IdApplicazione == 11 || _wb.IdApplicazione == 12)
                toUpdate = AggiornaData() || toUpdate;

            if (DataBase.OpenConnection())
            {
                InitUtente();
                DataBase.DB.SetParameters(_wb.DataAttiva.ToString("yyyyMMdd"), _wb.IdUtente, _wb.IdApplicazione);

                Workbook.AggiornaParametriApplicazione();

                if (Workbook.Repository.Applicazione != null)
                {
                    Simboli.rgbSfondo = Workbook.GetRGBFromString(Workbook.Repository.Applicazione["BackColorApp"].ToString());
                    Simboli.rgbTitolo = Workbook.GetRGBFromString(Workbook.Repository.Applicazione["BackColorFrameApp"].ToString());
                    Simboli.rgbLinee = Workbook.GetRGBFromString(Workbook.Repository.Applicazione["BorderColorApp"].ToString());
                }
                
                InitLog();

                Repository.DaAggiornare = toUpdate;

                return false;
            }
            else //Emergenza
            {
                if (toUpdate)
                {
                    System.Windows.Forms.MessageBox.Show("Il foglio non è inizializzato e non c'è connessione ad DB... Impossibile procedere! L'applicazione verrà chiusa.", "ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                    _wb.Base.Close();
                    return false;
                }

                DataBase.DB.SetParameters(_wb.DataAttiva.ToString("yyyyMMdd"), 0, 0);
                Simboli.nomeApplicazione = Workbook.Repository.Applicazione["DesApplicazione"].ToString();
                Struct.intervalloGiorni = Workbook.Repository.Applicazione["IntervalloGiorniEntita"] is DBNull ? 0 : int.Parse(Workbook.Repository.Applicazione["IntervalloGiorniEntita"].ToString());
                Struct.visualizzaRiepilogo = Workbook.Repository.Applicazione["VisRiepilogo"] is DBNull ? true : Workbook.Repository.Applicazione["VisRiepilogo"].Equals("1");

                return true;
            }
        }


        public static void Update()
        {
            //UPDATE
            string updatePath = Path.Combine(_wb.Path, "UPDATE");
            if (Directory.Exists(updatePath) && Directory.GetFiles(updatePath, _wb.Name).Any())
            {
                string name = _wb.Name;
                string fullName = _wb.FullName;
                _wb.Base.SaveAs(Path.Combine(_wb.Path, "old_" + _wb.Name));
                File.Copy(Path.Combine(updatePath, name), fullName, true);
                File.Delete(Path.Combine(updatePath, name));
                Application.Workbooks.Open(fullName);
                Simboli.Aborted = true;
                _wb.Base.Windows[1].Visible = false;
                return;
            }
            else
            {
                Window = new Win32Window(new IntPtr(Workbook.Application.Hwnd));
                try { if (File.Exists(Path.Combine(_wb.Path, "old_" + _wb.Name))) File.Delete(Path.Combine(_wb.Path, "old_" + _wb.Name)); }
                catch { }
            }
        }

        public static void StartUp(IToolsExcelThisWorkbook wb)
        {
            _wb = wb;

            Repository = new Utility.Repository(wb);
            Update();

            Application.ScreenUpdating = false;
            Application.Iteration = true;
            Application.MaxIterations = 100;
            Application.EnableEvents = false;

            Style.StdStyles();

            foreach (Excel.Worksheet ws in CategorySheets)
            {
                ws.Activate();
                ws.Range["A1"].Select();
                Application.ActiveWindow.ScrollRow = 1;
            }

            Main.Select();
            Application.WindowState = Excel.XlWindowState.xlMaximized;

            Simboli.pwd = AppSettings("pwd");

            bool wasProtected = Sheet.Protected;
            if (wasProtected)
                Sheet.Protected = false;

            Workbook.ScreenUpdating = false;

            bool emergenza = Initialize();

            Riepilogo r = new Riepilogo(Main);

            if (emergenza)
                r.RiepilogoInEmergenza();

            r.InitLabels();

            InsertLog(Core.DataBase.TipologiaLOG.LogAccesso, "Log on - " + Environment.UserName + " - " + Environment.MachineName);

            if (wasProtected)
                Sheet.Protected = true;
            Application.ScreenUpdating = true;
            Application.EnableEvents = true;
        }
        #endregion

        #region Log
        public static void InsertLog(Core.DataBase.TipologiaLOG logType, string message)
        {
            Excel.Worksheet log = _wb.Sheets["Log"];
            bool prot = log.ProtectContents;
            if (prot) log.Unprotect(Simboli.pwd);
            DataBase db = new DataBase();
            db.InsertLog(logType, message);
            if (prot) log.Protect(Simboli.pwd);
        }
        public static void RefreshLog()
        {
            Excel.Worksheet log = _wb.Sheets["Log"];
            bool prot = log.ProtectContents;
            if (prot) log.Unprotect(Simboli.pwd);
            DataBase db = new DataBase();
            db.RefreshLog();
            if (prot) log.Protect(Simboli.pwd);
        }
        #endregion

        #region Close
        public static void Save()
        {
            _wb.Base.Save();
        }
        public static void Close()
        {
            if (DataBase.LocalDB != null)
            {
                Simboli.EmergenzaForzata = false;
                Application.ScreenUpdating = false;
                if (WB.Application.DisplayDocumentActionTaskPane)
                    WB.Application.DisplayDocumentActionTaskPane = false;

                Main.Select();
                if (Simboli.ModificaDati)
                {
                    Sheet.Protected = false;
                    Simboli.ModificaDati = false;
                    Sheet.AbilitaModifica(false);
                    Sheet.SalvaModifiche();
                    Sheet.Protected = true;
                }
                DataBase.SalvaModificheDB();
                InsertLog(Core.DataBase.TipologiaLOG.LogAccesso, "Log off - " + Environment.UserName + " - " + Environment.MachineName);

                Application.ScreenUpdating = true;
            }
            Save();
        }
        #endregion

        #endregion
    }
}
