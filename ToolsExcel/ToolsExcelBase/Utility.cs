using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Core;
using Iren.ToolsExcel.UserConfig;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Deployment.Application;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Iren.ToolsExcel.Utility
{
    public class DataBase 
    {
        #region Costanti

        public const string NAME = "LocalDB";
        public struct SP
        {
            public const string APPLICAZIONE = "spApplicazioneProprieta",
                APPLICAZIONE_INFORMAZIONE_D = "spApplicazioneInformazioneD",
                APPLICAZIONE_INFORMAZIONE_H = "spApplicazioneInformazioneH",
                APPLICAZIONE_INFORMAZIONE_H_EXPORT = "PAR.spApplicazioneInformazioneH_Export",
                APPLICAZIONE_INFORMAZIONE_COMMENTO = "spApplicazioneInformazioneCommento",
                APPLICAZIONE_INIT = "spApplicazioneInit",
                APPLICAZIONE_LOG = "spApplicazioneLog2",
                APPLICAZIONE_NOTE = "spApplicazioneNote",
                APPLICAZIONE_RIBBON = "spApplicazioneRibbon",
                APPLICAZIONE_RIEPILOGO = "spApplicazioneRiepilogo",
                AZIONE = "spAzione",
                AZIONE_CATEGORIA = "spAzioneCategoria",
                CALCOLO = "spCalcolo",
                CALCOLO_INFORMAZIONE = "spCalcoloInformazione",
                CARICA_AZIONE_INFORMAZIONE = "spCaricaAzioneInformazione",
                CATEGORIA = "spCategoria",
                CATEGORIA_ENTITA = "spCategoriaEntita",
                CHECK_FONTE_METEO = "spCheckFonteMeteo",
                CHECKMODIFICASTRUTTURA = "spCheckModificaStruttura",
                ENTITA_ASSETTO = "spEntitaAssetto",
                ENTITA_AZIONE = "spEntitaAzione",
                ENTITA_AZIONE_CALCOLO = "spEntitaAzioneCalcolo",
                ENTITA_AZIONE_INFORMAZIONE = "spEntitaAzioneInformazione",
                ENTITA_CALCOLO = "spEntitaCalcolo",
                ENTITA_COMMITMENT = "spEntitaCommitment",
                ENTITA_GRAFICO = "spEntitaGrafico",
                ENTITA_GRAFICO_INFORMAZIONE = "spEntitaGraficoInformazione",
                ENTITA_INFORMAZIONE = "spEntitaInformazione",
                ENTITA_INFORMAZIONE_FORMATTAZIONE = "spEntitaInformazioneFormattazione",
                ENTITA_PARAMETRO_D = "spEntitaParametroD",
                ENTITA_PARAMETRO_H = "spEntitaParametroH",
                ENTITA_PROPRIETA = "spEntitaProprieta",
                ENTITA_RAMPA = "spEntitaRampa",
                GET_ORE_FERMATA = "spGetOreFermata",
                GET_VERSIONE = "spGetVersione",
                INSERT_APPLICAZIONE_INFORMAZIONE_XML = "spInsertApplicazioneInformazioneXML2",
                INSERT_APPLICAZIONE_RIEPILOGO = "spInsertApplicazioneRiepilogo",
                INSERT_LOG = "spInsertLog",
                INSERT_PROGRAMMAZIONE_PARAMETRO = "spInsertProgrammazione_Parametro",
                STAGIONE = "spTipologiaStagione",
                TIPOLOGIA_CHECK = "spTipologiaCheck",
                UTENTE = "spUtente";

            public struct PAR 
            {
                public const string DELETE_PARAMETRO = "PAR.spDeleteParametro",
                ELENCO_PARAMETRI = "PAR.spElencoParametri",
                INSERT_PARAMETRO = "PAR.spInsertParametro",
                VALORI_PARAMETRI = "PAR.spValoriParametri",
                UPDATE_PARAMETRO = "PAR.spUpdateParametro";
            }

            public struct RIBBON 
            {
                public const string GRUPPO_CONTROLLO = "RIBBON.spGruppoControllo",
                    CONTROLLO_APPLICAZIONE = "RIBBON.spControlloApplicazione",
                    CONTROLLO_FUNZIONE = "RIBBON.spControlloFunzione";
            }
                
        }
        public struct TAB
        {
            public const string ADDRESS_FROM = "AddressFrom",
                ADDRESS_TO = "AddressTo",
                ANNOTA = "AnnotaModifica",
                APPLICAZIONE_RIBBON = "ApplicazioneRibbon",
                AZIONE = "Azione",
                AZIONE_CATEGORIA = "AzioneCategoria",
                CALCOLO = "Calcolo",
                CALCOLO_INFORMAZIONE = "CalcoloInformazione",
                CATEGORIA = "Categoria",
                CATEGORIA_ENTITA = "CategoriaEntita",
                CHECK = "Check",
                DATE_DEFINITE = "DefinedDates",
                DATI_APPLICAZIONE_H = "DatiApplicazioneH",
                DATI_APPLICAZIONE_D = "DatiApplicazioneD",
                DATI_APPLICAZIONE_COMMENTO = "DatiApplicazioneCommento",
                EDITABILI = "Editabili",
                ENTITA_ASSETTO = "EntitaAssetto",
                ENTITA_AZIONE = "EntitaAzione",
                ENTITA_AZIONE_CALCOLO = "EntitaAzioneCalcolo",
                ENTITA_AZIONE_INFORMAZIONE = "EntitaAzioneInformazione",
                ENTITA_CALCOLO = "EntitaCalcolo",
                ENTITA_COMMITMENT = "EntitaCommitment",
                ENTITA_GRAFICO = "EntitaGrafico",
                ENTITA_GRAFICO_INFORMAZIONE = "EntitaGraficoInformazione",
                ENTITA_INFORMAZIONE = "EntitaInformazione",
                ENTITA_INFORMAZIONE_FORMATTAZIONE = "EntitaInformazioneFormattazione",
                ENTITA_PARAMETRO_D = "EntitaParametroD",
                ENTITA_PARAMETRO_H = "EntitaParametroH",
                ENTITA_PROPRIETA = "EntitaProprieta",
                ENTITA_RAMPA = "EntitaRampa",
                EXPORT_XML = "ExportXML",
                LISTA_APPLICAZIONI = "ListaApplicazioni",
                LOG = "Log",
                MERCATI = "Mercati",
                MODIFICA = "Modifica",
                NOMI_DEFINITI = "DefinedNames",
                SALVADB = "SaveDB",
                SELECTION = "Selection",
                STAGIONE = "Stagione",
                TIPOLOGIA_CHECK = "TipologiaCheck";

            public struct RIBBON
            {
                public const string GRUPPO_CONTROLLO = "GruppoControllo",
                    CONTROLLO_APPLICAZIONE = "ControlloApplicazione",
                    CONTROLLO_FUNZIONE = "ControlloFunzione";
            }
        }

        #endregion

        #region Variabili

        //protected static DataSet _localDB = null;
        protected static Core.DataBase _db = null;

        #endregion

        #region Proprietà statiche

        public static Dictionary<Core.DataBase.NomiDB, ConnectionState> StatoDB
        {
            get
            {
                if (Simboli.EmergenzaForzata)
                {
                    return new Dictionary<Core.DataBase.NomiDB, ConnectionState>() 
                    {
                        {Core.DataBase.NomiDB.SQLSERVER, ConnectionState.Closed},
                        {Core.DataBase.NomiDB.IMP, ConnectionState.Closed},
                        {Core.DataBase.NomiDB.ELSAG, ConnectionState.Closed}
                    };
                }

                return _db.StatoDB;
            }
        }
        public static System.Version Versione { get { return _db.GetCurrentV(); } }

        public static bool IsInitialized { get; private set; }


        public static int IdUtente { get { return _db.IdUtente; } set { _db.IdUtente = value; } }
        public static int IdApplicazione { get { return _db.IdApplicazione; } set { _db.IdApplicazione = value; } }
        public static DateTime DataAttiva { get { return _db.DataAttiva; } set { _db.DataAttiva = value; } }

        #endregion

        #region Metodi Statici

        //StatoDBChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        //{
        //    AggiornaLabelStatoDB();
        //}        

        public static void AddPropertyChanged(System.ComponentModel.PropertyChangedEventHandler d)
        {
            _db.PropertyChanged += d;
        }

        /// <summary>
        /// Inizializza il nuovo Core.DataBase collegato al dbName che rappresenta l'ambiente Prod|Test|Dev.
        /// </summary>
        /// <param name="dbName">Nome (corrisponde all'ambiente) del Database.</param>
        public static void CreateNew(string ambiente) 
        {
            if(_db == null || _db.Ambiente != ambiente)
                _db = new Core.DataBase(ambiente);

            IsInitialized = true;
        }
        /// <summary>
        /// Cambio ambiente tra Prod|Test|Prod.
        /// </summary>
        /// <param name="ambiente">Nuovo ambiente.</param>
        public static void SwitchEnvironment(string ambiente) 
        {
            if (_db.Ambiente != ambiente)
            {
                Workbook.Ambiente = ambiente;

                Delegate[] list = _db.GetPropertyChangedInvocationList();

                _db = new Core.DataBase(ambiente);
                _db.SetParameters(Workbook.DataAttiva, Workbook.IdUtente, Workbook.IdApplicazione);
                foreach (Delegate d in list)
                {
                    EventInfo ei = _db.GetType().GetEvent("PropertyChanged");
                    ei.AddEventHandler(_db, d);
                }
            }
        }
        /// <summary>
        /// Salva le modifiche effettuate ai fogli sul DataBase. Il processo consiste nella creazione di un file XML contenente tutte le righe della tabella di Modifica e successivo svuotamento della tabella stessa. Il processo richiede una connessione aperta. Diversamente le modifiche vengono salvate nella cartella di Emergenza dove, al momento della successiva chiamata al metodo, vengono reinviati al server in ordine cronologico.
        /// </summary>
        public static void SalvaModificheDB() 
        {
            if (Workbook.Repository != null)
            {
                //prendo la tabella di modifica e controllo se è nulla
                DataTable modifiche = Workbook.Repository[TAB.MODIFICA];
                if (modifiche != null && Workbook.IdUtente != 0)   //non invia se l'utente non è configurato... in ogni caso la tabella è vuota!!
                {
                    //tolgo il namespace che altrimenti aggiunge informazioni inutili al file da mandare al server
                    DataTable dt = modifiche.Copy();
                    dt.TableName = modifiche.TableName;
                    dt.Namespace = "";

                    //vari path per la funzione del salvataggio delle modifiche sul server
                    var path = Workbook.GetUsrConfigElement("pathExportModifiche");

                    //path del caricatore sul server
                    string cartellaRemota = path.Value;
                    //path della cartella di emergenza
                    string cartellaEmergenza = path.Emergenza;
                    //path della cartella di archivio in cui copiare i file in caso di esito positivo nel saltavaggio
                    string cartellaArchivio = path.Archivio;

                    string fileName = "";
                    //se la connessione è aperta (in emergenza forzata sarà sempre false) ed esiste la cartella del caricatore
                    if (OpenConnection() && Directory.Exists(cartellaRemota))
                    {
                        //metto in lavorazione i file nella cartella di emergenza
                        string[] fileEmergenza = Directory.GetFiles(cartellaEmergenza);
                        bool executed = false;
                        if (fileEmergenza.Length > 0)
                        {
                            if (System.Windows.Forms.MessageBox.Show("Sono presenti delle modifiche non ancora salvate sul DB. Procedere con il salvataggio? \n\nPremere Sì per inviare i dati al server, No per cancellare definitivamente le modifiche.", Simboli.nomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes)
                            {
                                //il nome file contiene la data, quindi li metto in ordine cronologico
                                Array.Sort<string>(fileEmergenza);
                                foreach (string file in fileEmergenza)
                                {
                                    File.Move(file, Path.Combine(cartellaRemota, file.Split('\\').Last()));

                                    executed = DataBase.Insert(SP.INSERT_APPLICAZIONE_INFORMAZIONE_XML, new QryParams() { { "@NomeFile", file.Split('\\').Last() } });
                                    if (executed)
                                    {
                                        if (!Directory.Exists(cartellaArchivio))
                                            Directory.CreateDirectory(cartellaArchivio);

                                        File.Move(Path.Combine(cartellaRemota, file.Split('\\').Last()), Path.Combine(cartellaArchivio, file.Split('\\').Last()));
                                    }
                                    else
                                    {
                                        System.Windows.Forms.MessageBox.Show("Il server ha restituito un errore nel salvataggio. Le modifiche rimarranno comunque salvate in locale.", Simboli.nomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                                    }
                                }
                            }
                            else
                            {
                                foreach (string file in fileEmergenza)
                                    File.Delete(file);
                            }
                        }

                        //controllo se la tabella è vuota
                        if (dt.Rows.Count == 0)
                            return;

                        //salvo le modifiche appena effettuate
                        fileName = Path.Combine(cartellaRemota, Simboli.nomeApplicazione.Replace(" ", "").ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml");
                        dt.WriteXml(fileName);//, XmlWriteMode.WriteSchema);

                        //se la query indica che il processo è andato a buon fine, sposto in archivio
                        executed = DataBase.Insert(SP.INSERT_APPLICAZIONE_INFORMAZIONE_XML, new QryParams() { { "@NomeFile", fileName.Split('\\').Last() } });
                        if (executed)
                        {
                            if (!Directory.Exists(cartellaArchivio))
                                Directory.CreateDirectory(cartellaArchivio);

                            File.Move(fileName, Path.Combine(cartellaArchivio, fileName.Split('\\').Last()));
                        }
                        else
                        {
                            fileName = Path.Combine(cartellaEmergenza, Simboli.nomeApplicazione.Replace(" ", "").ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml");
                            dt.WriteXml(fileName);//, XmlWriteMode.WriteSchema);

                            Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "Errore nel salvataggio delle modifiche. Il file è si trova in " + Environment.MachineName);

                            System.Windows.Forms.MessageBox.Show("Il server ha restituito un errore nel salvataggio. Le modifiche rimarranno comunque salvate in locale.", Simboli.nomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        if (dt.Rows.Count == 0)
                            return;

                        //metto le modifiche nella cartella emergenza
                        fileName = Path.Combine(cartellaEmergenza, Simboli.nomeApplicazione.Replace(" ", "").ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml");
                        try
                        {
                            dt.WriteXml(fileName, XmlWriteMode.WriteSchema);
                        }
                        catch (DirectoryNotFoundException)
                        {
                            Directory.CreateDirectory(cartellaEmergenza);
                            dt.WriteXml(fileName, XmlWriteMode.WriteSchema);
                        }

                        System.Windows.Forms.MessageBox.Show("A causa di problemi di rete le modifiche sono state salvate in locale", Simboli.nomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    }

                    //svuoto la tabella di modifiche
                    modifiche.Clear();
                }
            }
        }
        /// <summary>
        /// Aggiunge la riga di riepilogo al DB in modo da far evidenziare la casella nel foglio Main del Workbook.
        /// </summary>
        /// <param name="siglaEntita">La sigla dell'entità di cui aggiungere il riepilogo.</param>
        /// <param name="siglaAzione">La sigla dell'azione di cui aggiungere il riepilogo.</param>
        /// <param name="giorno">Il giorno in cui aggiungere il riepilogo.</param>
        /// <param name="presente">Se il dato collegato alla coppia Entità - Azione è presente o no nel DB.</param>
        public static void InsertApplicazioneRiepilogo(object siglaEntita, object siglaAzione, DateTime giorno, bool presente = true) 
        {
            try
            {
                if (OpenConnection())
                {
                    QryParams parameters = new QryParams() {
                    {"@SiglaEntita", siglaEntita},
                    {"@SiglaAzione", siglaAzione},
                    {"@Data", giorno.ToString("yyyyMMdd")},
                    {"@Presente", presente ? "1" : "0"}
                };
                    _db.Insert(DataBase.SP.INSERT_APPLICAZIONE_RIEPILOGO, parameters);
                }
            }
            catch (Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "InsertApplicazioneRiepilogo [" + giorno + ", " + siglaEntita + ", " + siglaAzione + "]: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// Inizializza i valori di default e imposta tutte le informazioni che devono essere "trascinate" dai giorni precedenti
        /// </summary>
        public static void ExecuteSPApplicazioneInit() 
        {
            SplashScreen.UpdateStatus("Inizializzazione valori di default");
            Select(SP.APPLICAZIONE_INIT);
        }
        /// <summary>
        /// Funzione per l'apertura della connessione che considera anche la presenza del flag di Emergenza Forzata.
        /// </summary>
        /// <returns>True se la connessione viene aperta, false altrimenti.</returns>
        public static bool OpenConnection() 
        {
            if (!Simboli.EmergenzaForzata)
                return _db.OpenConnection();

            return false;
        }
        /// <summary>
        /// Funzione per chiudere la connessione che considera anche la presenza del flag di Emergenza Forzata.
        /// </summary>
        /// <returns>True se la connessione viene chiusa, false altrimenti.</returns>
        public static bool CloseConnection() 
        {
            if (!Simboli.EmergenzaForzata)
                return _db.CloseConnection();

            return false;
        }
        /// <summary>
        /// Funzione per eseguire una stored procedure. Fa un "override" della funzione fornita da Core.DataBase che considera la presenza del flag di Emergenza Forzata.
        /// </summary>
        /// <param name="storedProcedure">Stored procedure.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <param name="timeout">Timeout del comando.</param>
        /// <returns>DataTable contenente il risultato della storedProcedure.</returns>
        public static DataTable Select(string storedProcedure, QryParams parameters, int timeout = 300) 
        {
            if (OpenConnection())
            {
                DataTable dt = _db.Select(storedProcedure, parameters, timeout);
                CloseConnection();
                
                return dt;
            }

            return null;
        }
        /// <summary>
        /// Funzione per eseguire una stored procedure. Fa un "override" della funzione fornita da Core.DataBase che considera la presenza del flag di Emergenza Forzata.
        /// </summary>
        /// <param name="storedProcedure">Stored procedure.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <param name="timeout">Timeout del comando.</param>
        /// <returns>DataTable contenente il risultato della storedProcedure.</returns>
        public static DataTable Select(string storedProcedure, String parameters, int timeout = 300) 
        {
            if (OpenConnection())
            {
                DataTable dt = _db.Select(storedProcedure, parameters, timeout);
                CloseConnection();

                return dt;
            }

            return null;
        }
        /// <summary>
        /// Funzione per eseguire una stored procedure. Fa un "override" della funzione fornita da Core.DataBase che considera la presenza del flag di Emergenza Forzata.
        /// </summary>
        /// <param name="storedProcedure">Stored procedure.</param>
        /// <param name="timeout">Timeout del comando.</param>
        /// <returns>DataTable contenente il risultato della storedProcedure.</returns>
        public static DataTable Select(string storedProcedure, int timeout = 300) 
        {
            if (OpenConnection())
            {
                DataTable dt = _db.Select(storedProcedure, timeout);
                CloseConnection();

                return dt;
            }

            return null;
        }
        /// <summary>
        /// Funzione per eseguire una stored procedure. Fa un "override" della funzione fornita da Core.DataBase che considera la presenza del flag di Emergenza Forzata.
        /// </summary>
        /// <param name="storedProcedure">Stored procedure.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <returns>True se il comando è andato a buon fine, false altrimenti.</returns>
        public static bool Insert(string storedProcedure, QryParams parameters, int timeout = 300) 
        {
            if (OpenConnection())
            {
                bool o = _db.Insert(storedProcedure, parameters, timeout);
                CloseConnection();
                return o;
            }
            return false;
        }
        public static bool Insert(string storedProcedure, QryParams parameters, out Dictionary<string, object> outParams, int timeout = 300) 
        {
            if (OpenConnection())
            {
                bool o = _db.Insert(storedProcedure, parameters, out outParams, timeout);
                CloseConnection();
                return o;
            }
            outParams = null;
            return false;
        }

        public static bool Delete(string storedProcedure, QryParams parameters, int timeout = 300) 
        {
            if (OpenConnection())
            {
                bool o = _db.Delete(storedProcedure, parameters, timeout);
                CloseConnection();
                return o;
            }
            return false;
        }
        public static bool Delete(string storedProcedure, string parameters, int timeout = 300) 
        {
            if (OpenConnection())
            {
                bool o = _db.Delete(storedProcedure, parameters, timeout);
                CloseConnection();
                return o;
            }
            return false;
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Inserisce una riga di Log.
        /// </summary>
        /// <param name="logType">Tipologia di Log.</param>
        /// <param name="message">Messaggio del Log.</param>
        public void InsertLog(Core.DataBase.TipologiaLOG logType, string message)
        {
#if (!DEBUG)
            if (OpenConnection())
            {
                Insert(SP.INSERT_LOG, new QryParams() { { "@IdTipologia", logType }, { "@Messaggio", message } });
            }
#endif
            RefreshLog();
        }
        /// <summary>
        /// Aggiorna il foglio di Log.
        /// </summary>
        public void RefreshLog()
        {
            if (OpenConnection())
            {
                DataTable dt = Select(SP.APPLICAZIONE_LOG);
                if (dt != null)
                {
                    Workbook.LogDataTable.Clear();
                    Workbook.LogDataTable.Merge(dt);

                    if (Workbook.Log.ListObjects.Count > 0)
                        Workbook.Log.ListObjects[1].Range.EntireColumn.AutoFit();
                }
            }
        }        

        #endregion
    }

    public class Date 
    {
        #region Proprietà

        /// <summary>
        /// Scorciatoia per ottenere il suffisso della dataAttiva.
        /// </summary>
        public static string SuffissoDATA1
        {
            get { return GetSuffissoData(Workbook.DataAttiva); }
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Restituisce le ore di intervallo tra la data attiva e la data fine specificata.
        /// </summary>
        /// <param name="fine">Data fine.</param>
        /// <returns>Ore di intervallo.</returns>
        public static int GetOreIntervallo(DateTime fine)
        {
            return GetOreIntervallo(Workbook.DataAttiva, fine);
        }
        /// <summary>
        /// Restituisce le ore di intervallo tra una data inizio e fine specificate.
        /// </summary>
        /// <param name="inizio">Data inizio.</param>
        /// <param name="fine">Data fine.</param>
        /// <returns>Ore di intervallo.</returns>
        public static int GetOreIntervallo(DateTime inizio, DateTime fine)
        {
            return (int)(fine.AddDays(1).ToUniversalTime() - inizio.ToUniversalTime()).TotalHours;
        }
        /// <summary>
        /// Restituisce le ore che compongono il giorno passato per parametro.
        /// </summary>
        /// <param name="giorno">Giorno.</param>
        /// <returns>Numero di ore del giorno.</returns>
        public static int GetOreGiorno(DateTime giorno)
        {
            DateTime giornoSucc = giorno.AddDays(1);
            return (int)(giornoSucc.ToUniversalTime() - giorno.ToUniversalTime()).TotalHours;
        }
        /// <summary>
        /// Restituisce le ore che compongono il giorno passato per parametro.
        /// </summary>
        /// <param name="suffissoData">Suffisso del giorno.</param>
        /// <returns>Numero di ore del giorno.</returns>
        public static int GetOreGiorno(string suffissoData)
        {
            return GetOreGiorno(GetDataFromSuffisso(suffissoData));
        }
        /// <summary>
        /// Restituisce il suffisso del giorno rispetto alla data attiva.
        /// </summary>
        /// <param name="giorno">Giorno di cui trovare il suffisso.</param>
        /// <returns>Stringa del tipo DATAx con x = 1 se giorno è data attiva, x = 2 se giorno è data attiva + 1, e così via.</returns>
        public static string GetSuffissoData(DateTime giorno)
        {
            return GetSuffissoData(Utility.Workbook.DataAttiva, giorno);
        }
        /// <summary>
        /// Restituisce il suffisso del giorno rispetto alla data attiva.
        /// </summary>
        /// <param name="giorno">Giorno di cui trovare il suffisso.</param>
        /// <returns>Stringa del tipo DATAx con x = 1 se giorno è data attiva, x = 2 se giorno è data attiva + 1, e così via.</returns>
        public static string GetSuffissoData(string giorno)
        {
            return GetSuffissoData(Utility.Workbook.DataAttiva, giorno);
        }
        /// <summary>
        /// Restituisce il suffisso del giorno rispetto alla data inizio.
        /// </summary>
        /// <param name="inizio">Data di inizio.</param>
        /// <param name="giorno">Giorno di cui trovare il suffisso.</param>
        /// <returns>Stringa del tipo DATAx con x = 1 se giorno è data attiva, x = 2 se giorno è data attiva + 1, e così via.</returns>
        public static string GetSuffissoData(DateTime inizio, DateTime giorno)
        {
            if (inizio > giorno)
            {
                return "DATA0";
            }
            TimeSpan dayDiff = giorno - inizio;
            return "DATA" + (dayDiff.Days + 1);
        }
        /// <summary>
        /// Restituisce il suffisso del giorno rispetto alla data inizio.
        /// </summary>
        /// <param name="inizio">Data di inizio.</param>
        /// <param name="giorno">Giorno di cui trovare il suffisso.</param>
        /// <returns>Stringa del tipo DATAx con x = 1 se giorno è data attiva, x = 2 se giorno è data attiva + 1, e così via.</returns>
        public static string GetSuffissoData(DateTime inizio, object giorno)
        {
            DateTime day = DateTime.ParseExact(giorno.ToString().Substring(0, 8), "yyyyMMdd", CultureInfo.InvariantCulture);
            return GetSuffissoData(inizio, day);
        }
        /// <summary>
        /// Restituisce il suffisso dell'ora in ingresso.
        /// </summary>
        /// <param name="ora">Numero rappresentante l'ora da 1 a 25.</param>
        /// <returns>Stringa del tipo Hx con x = ora.</returns>
        public static string GetSuffissoOra(int ora)
        {
            return "H" + ora;
        }
        /// <summary>
        /// Restituisce il suffisso dell'ora estraendolo dalla data ISO yyyyMMddHH
        /// </summary>
        /// <param name="dataOra">Stringa nella forma DATAx.Hy.</param>
        /// <returns>Stringa del tipo Hx.</returns>
        public static string GetSuffissoOra(object dataOra)
        {
            string dtO = dataOra.ToString();
            if (dtO.Length != 10)
                return "";

            return GetSuffissoOra(int.Parse(dtO.Substring(dtO.Length - 2, 2)));
        }
        /// <summary>
        /// Restituisce la data in formato ISO yyyyMMddHH a partire dal suffisso data e suffisso ora.
        /// </summary>
        /// <param name="data">Suffisso data.</param>
        /// <param name="ora">Suffisso ora.</param>
        /// <returns>Data in formato ISO yyyyMMddHH.</returns>
        public static string GetDataFromSuffisso(string data, string ora)
        {
            DateTime outDate = GetDataFromSuffisso(data);
            ora = ora == "" ? "0" : ora;
            int outOra = int.Parse(Regex.Match(ora, @"\d+").Value);

            return outDate.ToString("yyyyMMdd") + (outOra != 0 ? outOra.ToString("D2") : "");
        }
        /// <summary>
        /// Restituisce la data in formato ISO yyyyMMdd a partire dal suffisso data.
        /// </summary>
        /// <param name="data">Suffisso data.</param>
        /// <returns>Data in formato ISO yyyyMMdd.</returns>
        public static DateTime GetDataFromSuffisso(string data)
        {
            int giorno = int.Parse(Regex.Match(data.ToString(), @"\d+").Value);
            return Workbook.DataAttiva.AddDays(giorno - 1);
        }
        /// <summary>
        /// Restituisce l'ora a partire dalla stringa in formato ISO yyyyMMddHH
        /// </summary>
        /// <param name="dataOra"></param>
        /// <returns></returns>
        public static int GetOraFromDataOra(string dataOra)
        {
            string dtO = dataOra.ToString();
            if (dtO.Length != 10)
                return -1;

            return int.Parse(dtO.Substring(dtO.Length - 2, 2));
        }
        /// <summary>
        /// Restituisce l'ora a partire dal suffisso ora del tipo Hx.
        /// </summary>
        /// <param name="suffissoOra">Suffisso ora.</param>
        /// <returns>Intero rappresentante l'ora (1 - 25).</returns>
        public static int GetOraFromSuffissoOra(string suffissoOra)
        {
            string match = Regex.Match(suffissoOra, @"\d+").Value;
            return int.Parse(match);
        }
        
        #endregion
    }

    class Win32Window : IWin32Window
    {
        public Win32Window(IntPtr handle) { Handle = handle; }
        public IntPtr Handle { get; private set; }
    }
}
