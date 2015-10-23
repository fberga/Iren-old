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
                TIPOLOGIA_CHECK = "spTipologiaCheck",
                UTENTE = "spUtente",

                DELETE_PARAMETRO = "PAR.spDeleteParametro",
                ELENCO_PARAMETRI = "PAR.spElencoParametri",
                INSERT_PARAMETRO = "PAR.spInsertParametro",
                VALORI_PARAMETRI = "PAR.spValoriParametri",
                UPDATE_PARAMETRO = "PAR.spUpdateParametro";
        }
        public struct Tab
        {
            public const string ADDRESS_FROM = "AddressFrom",
                ADDRESS_TO = "AddressTo",
                ANNOTA = "AnnotaModifica",
                //APPLICAZIONE = "Applicazione",
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
                MODIFICA = "Modifica",
                NOMI_DEFINITI = "DefinedNames",
                SALVADB = "SaveDB",
                SELECTION = "Selection",
                TIPOLOGIA_CHECK = "TipologiaCheck",
                UTENTE = "Utente";
        }

        #endregion

        #region Variabili

        protected static DataSet _localDB = null;
        protected static Core.DataBase _db = null;

        #endregion

        #region Proprietà

        public static DataSet LocalDB { get { return _localDB; } }
        public static Core.DataBase DB { get { return _db; } }
        public static DateTime DataAttiva { get { return _db.DataAttiva; } }
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

                return DB.StatoDB;
            }
        }

        #endregion

        #region Metodi Statici

        /// <summary>
        /// Inizializza il nuovo Core.DataBase collegato al dbName che rappresenta l'ambiente Prod|Test|Dev.
        /// </summary>
        /// <param name="dbName">Nome (corrisponde all'ambiente) del Database.</param>
        public static void Initialize(string dbName) 
        {
            _db = new Core.DataBase(dbName);
        }
        /// <summary>
        /// Inizializza il nuovo dataset Locale.
        /// </summary>
        //public static void InitNewLocalDB()
        //{


        //    _localDB = new DataSet(NAME);
        //}
        /// <summary>
        /// Se il dataset contiente la tabella identificata da name, la cancella. Va cancellata perché se nelle varie modifiche al DB viene cambiata la struttura, fare il merge comporterebbe il verificarsi di un errore di mancata corrispondenza delle strutture.
        /// </summary>
        /// <param name="name"></param>
        public static void ResetTable(string name) 
        {
            if (_localDB.Tables.Contains(name))
            {
                _localDB.Tables.Remove(name);
            }
        }
        /// <summary>
        /// Cambio l'ID applicazione su cui Core.DataBase basa le sue ricerche.
        /// </summary>
        /// <param name="appID">Nuovo ID applicazione.</param>
        public static void ChangeAppID(string appID)
        {
            Workbook.ChangeAppSettings("AppID", appID);
            _db.ChangeAppID(int.Parse(appID));
        }
        /// <summary>
        /// Cambio la Data Attiva su cui Core.DataBase basa le sue ricerche.
        /// </summary>
        /// <param name="dataAttiva">Nuova DataAttiva.</param>
        public static void ChangeDate(DateTime dataAttiva) 
        {
            _db.ChangeDate(dataAttiva.ToString("yyyyMMdd"));
        }
        /// <summary>
        /// Cambio ambiente tra Prod|Test|Prod.
        /// </summary>
        /// <param name="ambiente">Nuovo ambiente.</param>
        public static void SwitchEnvironment(string ambiente) 
        {            
            Workbook.ChangeAppSettings("DB", ambiente);

            int idA = _db.IdApplicazione;
            int idU = _db.IdUtenteAttivo;
            string data = _db.DataAttiva.ToString("yyyyMMdd");

            _db = new Core.DataBase(ambiente);
            _db.SetParameters(data, idU, idA);
        }
        /// <summary>
        /// Salva le modifiche effettuate ai fogli sul DataBase. Il processo consiste nella creazione di un file XML contenente tutte le righe della tabella di Modifica e successivo svuotamento della tabella stessa. Il processo richiede una connessione aperta. Diversamente le modifiche vengono salvate nella cartella di Emergenza dove, al momento della successiva chiamata al metodo, vengono reinviati al server in ordine cronologico.
        /// </summary>
        public static void SalvaModificheDB() 
        {
            if (LocalDB != null)
            {
                //prendo la tabella di modifica e controllo se è nulla
                DataTable modifiche = LocalDB.Tables[Tab.MODIFICA];
                if (modifiche != null && DataBase.DB.IdUtenteAttivo != 0)   //non invia se l'utente non è configurato... in ogni caso la tabella è vuota!!
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

                                    executed = DataBase.DB.Insert(SP.INSERT_APPLICAZIONE_INFORMAZIONE_XML, new QryParams() { { "@NomeFile", file.Split('\\').Last() } });
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
                        executed = DataBase.DB.Insert(SP.INSERT_APPLICAZIONE_INFORMAZIONE_XML, new QryParams() { { "@NomeFile", fileName.Split('\\').Last() } });
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
            //non posso cancellare la tabella altrimenti perdo il dataBinding con l'oggetto nel foglio di Log. Quindi svuoto la tabella e la riempio con i nuovi valori.
            if (OpenConnection())
            {
                DataTable dt = Select(SP.APPLICAZIONE_LOG);
                if (dt != null)
                {
                    dt.TableName = Tab.LOG;

                    bool sameSchema = dt.Columns.Count == _localDB.Tables[Tab.LOG].Columns.Count;

                    for (int i = 0; i < dt.Columns.Count && sameSchema; i++)
                        if (_localDB.Tables[Tab.LOG].Columns[i].ColumnName != dt.Columns[i].ColumnName)
                            sameSchema = false;

                    //svuoto la tabella allo stato attuale
                    _localDB.Tables[Tab.LOG].Clear();
                    //la riempio con tutte le rige comprese le nuove
                    _localDB.Tables[Tab.LOG].Merge(dt);

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
            get { return GetSuffissoData(DataBase.DataAttiva); }
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
            return GetOreIntervallo(DataBase.DataAttiva, fine);
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
            return GetSuffissoData(Utility.DataBase.DataAttiva, giorno);
        }
        /// <summary>
        /// Restituisce il suffisso del giorno rispetto alla data attiva.
        /// </summary>
        /// <param name="giorno">Giorno di cui trovare il suffisso.</param>
        /// <returns>Stringa del tipo DATAx con x = 1 se giorno è data attiva, x = 2 se giorno è data attiva + 1, e così via.</returns>
        public static string GetSuffissoData(string giorno)
        {
            return GetSuffissoData(Utility.DataBase.DataAttiva, giorno);
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
            return DataBase.DB.DataAttiva.AddDays(giorno - 1);
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

    //public class Repository : DataBase
    //{
    //    #region Variabili

    //    private static bool _daAggiornare = false;
    //    private static bool _isMultiApplication = false;
    //    private static string[] _appIDs = null;

    //    #endregion

    //    #region Proprietà

    //    public static bool DaAggiornare { get { return _daAggiornare; } set { _daAggiornare = value; } }
    //    public static bool IsMultiApplicaion { get { return _isMultiApplication; } }

    //    #endregion

    //    #region Metodi

    //    public static void InitStrutturaNomi()
    //    {
    //        CreaTabellaNomi();
    //        CreaTabellaDate();
    //        CreaTabellaAddressFrom();
    //        CreaTabellaAddressTo();
    //        CreaTabellaModifica();
    //        CreaTabellaExportXML();
    //        CreaTabellaEditabili();
    //        CreaTabellaSalvaDB();
    //        CreaTabellaAnnotaModifica();
    //        CreaTabellaCheck();
    //        CreaTabellaSelection();
    //        _localDB.AcceptChanges();
    //    }

    //    /// <summary>
    //    /// Launcher per tutte le funzioni che aggiornano il repository in seguito alla richiesta di aggiornare la struttura.
    //    /// </summary>
    //    public static void Aggiorna(string[] appIDs)
    //    {
    //        _isMultiApplication = appIDs != null;
    //        _appIDs = appIDs;

    //        InitStrutturaNomi();
    //        CaricaApplicazioni();
    //        CaricaAzioni();
    //        CaricaCategorie();
    //        CaricaApplicazioneRibbon();
    //        CaricaAzioneCategoria();
    //        CaricaCategoriaEntita();
    //        CaricaEntitaAzione();
    //        CaricaEntitaAzioneCalcolo();
    //        CaricaEntitaInformazione();
    //        CaricaEntitaAzioneInformazione();
    //        CaricaCalcolo();
    //        CaricaCalcoloInformazione();
    //        CaricaEntitaCalcolo();
    //        CaricaEntitaGrafico();
    //        CaricaEntitaGraficoInformazione();
    //        CaricaEntitaCommitment();
    //        CaricaEntitaRampa();
    //        CaricaEntitaAssetto();
    //        CaricaEntitaProprieta();
    //        CaricaEntitaInformazioneFormattazione();
    //        CaricaEntitaParametroD();
    //        CaricaEntitaParametroH();
            
    //        _localDB.AcceptChanges();
    //    }
    //    #region Aggiorna Struttura Dati

    //    #region Init Struttura Nomi
    //    private static bool CreaTabellaNomi()
    //    {
    //        try
    //        {
    //            string name = Tab.NOMI_DEFINITI;
    //            ResetTable(name);
    //            DataTable dt = DefinedNames.GetDefaultNameTable(name);
    //            _localDB.Tables.Add(dt);
    //            return true;
    //        }
    //        catch (Exception)
    //        {
    //            return false;
    //        }
    //    }
    //    private static bool CreaTabellaDate()
    //    {
    //        try
    //        {
    //            string name = Tab.DATE_DEFINITE;
    //            ResetTable(name);
    //            DataTable dt = DefinedNames.GetDefaultDateTable(name);
    //            _localDB.Tables.Add(dt);
    //            return true;
    //        }
    //        catch (Exception)
    //        {
    //            return false;
    //        }
    //    }
    //    private static bool CreaTabellaAddressFrom()
    //    {
    //        try
    //        {
    //            string name = Tab.ADDRESS_FROM;
    //            ResetTable(name);
    //            DataTable dt = DefinedNames.GetDefaultAddressFromTable(name);
    //            _localDB.Tables.Add(dt);
    //            return true;
    //        }
    //        catch (Exception)
    //        {
    //            return false;
    //        }
    //    }
    //    private static bool CreaTabellaAddressTo()
    //    {
    //        try
    //        {
    //            string name = Tab.ADDRESS_TO;
    //            ResetTable(name);
    //            DataTable dt = DefinedNames.GetDefaultAddressToTable(name);
    //            _localDB.Tables.Add(dt);
    //            return true;
    //        }
    //        catch (Exception)
    //        {
    //            return false;
    //        }
    //    }
    //    private static bool CreaTabellaEditabili()
    //    {
    //        try
    //        {
    //            string name = Tab.EDITABILI;
    //            ResetTable(name);
    //            DataTable dt = DefinedNames.GetDefaultEditableTable(name);
    //            _localDB.Tables.Add(dt);
    //            return true;
    //        }
    //        catch (Exception)
    //        {
    //            return false;
    //        }
    //    }
    //    private static bool CreaTabellaSalvaDB()
    //    {
    //        try
    //        {
    //            string name = Tab.SALVADB;
    //            ResetTable(name);
    //            DataTable dt = DefinedNames.GetDefaultSaveTable(name);
    //            _localDB.Tables.Add(dt);
    //            return true;
    //        }
    //        catch (Exception)
    //        {
    //            return false;
    //        }
    //    }
    //    private static bool CreaTabellaAnnotaModifica()
    //    {
    //        try
    //        {
    //            string name = Tab.ANNOTA;
    //            ResetTable(name);
    //            DataTable dt = DefinedNames.GetDefaultToNoteTable(name);
    //            _localDB.Tables.Add(dt);
    //            return true;
    //        }
    //        catch (Exception)
    //        {
    //            return false;
    //        }
    //    }
    //    private static bool CreaTabellaModifica()
    //    {
    //        try
    //        {
    //            string name = Tab.MODIFICA;
    //            ResetTable(name);
    //            DataTable dt = new DataTable(name)
    //            {
    //                Columns =
    //                {
    //                    {"SiglaEntita", typeof(string)},
    //                    {"SiglaInformazione", typeof(string)},
    //                    {"Data", typeof(string)},
    //                    {"Valore", typeof(string)},
    //                    {"AnnotaModifica", typeof(string)},
    //                    {"IdApplicazione", typeof(string)},
    //                    {"IdUtente", typeof(string)}
    //                }
    //            };

    //            dt.PrimaryKey = new DataColumn[] { dt.Columns["SiglaEntita"], dt.Columns["SiglaInformazione"], dt.Columns["Data"] };
    //            _localDB.Tables.Add(dt);
    //            return true;
    //        }
    //        catch (Exception)
    //        {
    //            return false;
    //        }
    //    }
    //    private static bool CreaTabellaExportXML()
    //    {
    //        try
    //        {
    //            string name = Tab.EXPORT_XML;
    //            ResetTable(name);
    //            DataTable dt = new DataTable(name)
    //            {
    //                Columns =
    //                {
    //                    {"SiglaEntita", typeof(string)},
    //                    {"SiglaInformazione", typeof(string)},
    //                    {"Data", typeof(string)},
    //                    {"Valore", typeof(string)},
    //                    {"AnnotaModifica", typeof(string)},
    //                    {"IdApplicazione", typeof(string)},
    //                    {"IdUtente", typeof(string)}
    //                }
    //            };

    //            dt.PrimaryKey = new DataColumn[] { dt.Columns["SiglaEntita"], dt.Columns["SiglaInformazione"], dt.Columns["Data"] };
    //            _localDB.Tables.Add(dt);
    //            return true;
    //        }
    //        catch (Exception)
    //        {
    //            return false;
    //        }
    //    }
    //    private static bool CreaTabellaCheck()
    //    {
    //        try
    //        {
    //            string name = Tab.CHECK;
    //            ResetTable(name);
    //            DataTable dt = DefinedNames.GetDefaultCheckTable(name);
    //            _localDB.Tables.Add(dt);
    //            return true;
    //        }
    //        catch (Exception)
    //        {
    //            return false;
    //        }
    //    }
    //    private static bool CreaTabellaSelection()
    //    {
    //        try
    //        {
    //            string name = Tab.SELECTION;
    //            ResetTable(name);
    //            DataTable dt = DefinedNames.GetDefaultSelectionTable(name);
    //            _localDB.Tables.Add(dt);
    //            return true;
    //        }
    //        catch (Exception)
    //        {
    //            return false;
    //        }
    //    }
    //    #endregion

    //    /// <summary>
    //    /// Metodo richiamato da tutte le routine sottostanti che effettua la chiamata alla stored procedure sul server e aggiunge la tabella al DataSet locale. Restituisce true se l'operazione è andata a buon fine, lancia un'eccezione RepositoryUpdateException se fallisce.
    //    /// </summary>
    //    /// <param name="tableName">Nome della tabella da aggiornare.</param>
    //    /// <param name="spName">Nome della stored procedure da eseguire.</param>
    //    /// <param name="parameters">Parametri della stored procedure.</param>
    //    /// <returns>True se l'operazione è andata a buon fine.</returns>
    //    public static void CaricaDati(string tableName, string spName, QryParams parameters)
    //    {
    //        DataTable dt = new DataTable();
    //        if (_isMultiApplication)
    //        {
    //            foreach (string id in _appIDs)
    //            {
    //                parameters["@IdApplicazione"] = id;
    //                dt.Merge(Select(spName, parameters) ?? new DataTable());
    //            }
    //            if (dt.Columns.Count == 0)
    //                dt = null;
    //        }
    //        else
    //        {
    //            dt = Select(spName, parameters);
    //        }
    //        if (dt != null)
    //        {
    //            ResetTable(tableName);
    //            dt.TableName = tableName;
    //            _localDB.Tables.Add(dt);
    //        }
    //    }
    //    /// <summary>
    //    /// Carica la lista di tutte le applicazioni disponibili.
    //    /// </summary>
    //    public static void CaricaApplicazioni()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@IdApplicazione", 0}
    //            };

    //        CaricaDati(Tab.LISTA_APPLICAZIONI, SP.APPLICAZIONE, parameters);
    //    }
    //    /// <summary>
    //    /// Carica i dati necessari alla creazione del menu ribbon.
    //    /// </summary>
    //    /// <returns></returns>
    //    public static void CaricaApplicazioneRibbon()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@IdApplicazione", DB.IdApplicazione},
    //                {"@IdUtente", DB.IdUtenteAttivo}
    //            };

    //        CaricaDati(Tab.APPLICAZIONE_RIBBON, SP.APPLICAZIONE_RIBBON, parameters);
    //    }
    //    /// <summary>
    //    /// Carica le azioni.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaAzioni()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaAzione", Core.DataBase.ALL},
    //                {"@Operativa", Core.DataBase.ALL},
    //                {"@Visibile", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.AZIONE, SP.AZIONE, parameters);
    //    }
    //    /// <summary>
    //    /// Carica le categorie.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaCategorie()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaCategoria", Core.DataBase.ALL},
    //                {"@Operativa", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.CATEGORIA, SP.CATEGORIA, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione azione categoria.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaAzioneCategoria()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaAzione", Core.DataBase.ALL},
    //                {"@SiglaCategoria", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.AZIONE_CATEGORIA, SP.AZIONE_CATEGORIA, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione categoria entita.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaCategoriaEntita()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaCategoria", Core.DataBase.ALL},
    //                {"@SiglaEntita", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.CATEGORIA_ENTITA, SP.CATEGORIA_ENTITA, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità azione.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaAzione()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaEntita", Core.DataBase.ALL},
    //                {"@SiglaAzione", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.ENTITA_AZIONE, SP.ENTITA_AZIONE, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità azione calcolo.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaAzioneCalcolo()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaEntita", Core.DataBase.ALL},
    //                {"@SiglaAzione", Core.DataBase.ALL},
    //                {"@SiglaCalcolo", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.ENTITA_AZIONE_CALCOLO, SP.ENTITA_AZIONE_CALCOLO, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità informazione.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaInformazione()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaEntita", Core.DataBase.ALL},
    //                {"@SiglaInformazione", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.ENTITA_INFORMAZIONE, SP.ENTITA_INFORMAZIONE, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità azione informazione.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaAzioneInformazione()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaEntita", Core.DataBase.ALL},
    //                {"@SiglaAzione", Core.DataBase.ALL},
    //                {"@SiglaInformazione", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.ENTITA_AZIONE_INFORMAZIONE, SP.ENTITA_AZIONE_INFORMAZIONE, parameters);
    //    }
    //    /// <summary>
    //    /// Carica i calcoli.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaCalcolo()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaCalcolo", Core.DataBase.ALL},
    //                {"@IdTipologiaCalcolo", 0}
    //            };

    //        CaricaDati(Tab.CALCOLO, SP.CALCOLO, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione calcolo informazione.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaCalcoloInformazione()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaCalcolo", Core.DataBase.ALL},
    //                {"@SiglaInformazione", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.CALCOLO_INFORMAZIONE, SP.CALCOLO_INFORMAZIONE, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità calcolo.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaCalcolo()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaEntita", Core.DataBase.ALL},
    //                {"@SiglaCalcolo", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.ENTITA_CALCOLO, SP.ENTITA_CALCOLO, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità grafico.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaGrafico()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaEntita", Core.DataBase.ALL},
    //                {"@SiglaGrafico", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.ENTITA_GRAFICO, SP.ENTITA_GRAFICO, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità grafico informazione.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaGraficoInformazione()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaEntita", Core.DataBase.ALL},
    //                {"@SiglaGrafico", Core.DataBase.ALL},
    //                {"@SiglaInformazione", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.ENTITA_GRAFICO_INFORMAZIONE, SP.ENTITA_GRAFICO_INFORMAZIONE, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità commitment.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaCommitment()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaEntita", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.ENTITA_COMMITMENT, SP.ENTITA_COMMITMENT, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità rampa.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaRampa()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaEntita", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.ENTITA_RAMPA, SP.ENTITA_RAMPA, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità assetto.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaAssetto()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaEntita", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.ENTITA_ASSETTO, SP.ENTITA_ASSETTO, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità proprietà.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaProprieta()
    //    {
    //        CaricaDati(Tab.ENTITA_PROPRIETA, SP.ENTITA_PROPRIETA, new QryParams());
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità informazione formattazione.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaInformazioneFormattazione()
    //    {
    //        QryParams parameters = new QryParams() 
    //            {
    //                {"@SiglaEntita", Core.DataBase.ALL},
    //                {"@SiglaInformazione", Core.DataBase.ALL}
    //            };

    //        CaricaDati(Tab.ENTITA_INFORMAZIONE_FORMATTAZIONE, SP.ENTITA_INFORMAZIONE_FORMATTAZIONE, parameters);
    //    }
    //    /// <summary>
    //    /// Carica la tipologia check.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaTipologiaCheck()
    //    {
    //        CaricaDati(Tab.TIPOLOGIA_CHECK, SP.TIPOLOGIA_CHECK, new QryParams());
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità parametro giornaliero.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaParametroD()
    //    {
    //        CaricaDati(Tab.ENTITA_PARAMETRO_D, SP.ENTITA_PARAMETRO_D, new QryParams());
    //    }
    //    /// <summary>
    //    /// Carica la relazione entità parametro orario.
    //    /// </summary>
    //    /// <returns></returns>
    //    private static void CaricaEntitaParametroH()
    //    {
    //        CaricaDati(Tab.ENTITA_PARAMETRO_H, SP.ENTITA_PARAMETRO_H, new QryParams());
    //    }

    //    #endregion

    //    /// <summary>
    //    /// Carica solo i parametri dell'applicazione contenuti nella tabella APPLICAZIONE.
    //    /// </summary>
    //    /// <param name="appID">L'ID dell'applicazione di cui ricaricare i parametri.</param>
    //    /// <returns>La tabella contenente tutti i dati trovati sull'applicazione.</returns>
    //    public static DataTable CaricaApplicazione(object appID)
    //    {
    //        string name = DataBase.Tab.APPLICAZIONE;
    //        DataBase.ResetTable(name);
    //        DataTable dt = DataBase.Select(DataBase.SP.APPLICAZIONE) ?? new DataTable();
            
    //        dt.TableName = name;
    //        return dt;
    //    }

    //    #endregion
    //}

    public class Repository : IEnumerable
    {
        #region Variabili

        private IToolsExcelThisWorkbook _wb;
        private static bool _isMultiApplication = false;
        private static string[] _appIDs = null;

        #endregion

        #region Proprietà

        public DataTable this[string tableName] 
        { 
            get 
            {
                if (_wb.RepositoryDataSet.Tables.Contains(tableName))
                    return _wb.RepositoryDataSet.Tables[tableName];

                return null;
            }
            private set
            {
                if (_wb.RepositoryDataSet.Tables.Contains(tableName))
                    _wb.RepositoryDataSet.Tables.Remove(tableName);
                if (value.TableName != tableName)
                    value.TableName = tableName;
                _wb.RepositoryDataSet.Tables.Add(value);
            }
        }
        public DataTable this[int index]
        {
            get
            {
                if (_wb.RepositoryDataSet.Tables.Count > index)
                    return _wb.RepositoryDataSet.Tables[index];

                return null;
            }
        }

        public DataRow Applicazione { get; private set; }

        public int TablesCount { get { return _wb.RepositoryDataSet.Tables.Count; } }

        public bool DaAggiornare { get; set; }

        #endregion


        public Repository(IToolsExcelThisWorkbook wb)
        {
            _wb = wb;
            DaAggiornare = false;
            Applicazione = null;
        }

        #region Metodi

        public void Aggiorna(string[] appIDs)
        {
            //_isMultiApplication = appIDs != null;
            //_appIDs = appIDs;

            InitStrutturaNomi();
            CaricaApplicazioni();
            CaricaAzioni();
            CaricaCategorie();
            CaricaApplicazioneRibbon();
            CaricaAzioneCategoria();
            CaricaCategoriaEntita();
            CaricaEntitaAzione();
            CaricaEntitaAzioneCalcolo();
            CaricaEntitaInformazione();
            CaricaEntitaAzioneInformazione();
            CaricaCalcolo();
            CaricaCalcoloInformazione();
            CaricaEntitaCalcolo();
            CaricaEntitaGrafico();
            CaricaEntitaGraficoInformazione();
            CaricaEntitaCommitment();
            CaricaEntitaRampa();
            CaricaEntitaAssetto();
            CaricaEntitaProprieta();
            CaricaEntitaInformazioneFormattazione();
            CaricaEntitaParametroD();
            CaricaEntitaParametroH();

            //_wb.RepositoryDataSet.AcceptChanges();
        }
        #region Aggiorna Struttura Dati

        #region Init Struttura Nomi
        
        public void InitStrutturaNomi()
        {

            this[DataBase.Tab.NOMI_DEFINITI] = DefinedNames.GetDefaultNameTable(DataBase.Tab.NOMI_DEFINITI);
            this[DataBase.Tab.DATE_DEFINITE] = DefinedNames.GetDefaultDateTable(DataBase.Tab.DATE_DEFINITE);
            this[DataBase.Tab.ADDRESS_FROM] = DefinedNames.GetDefaultAddressFromTable(DataBase.Tab.ADDRESS_FROM);
            this[DataBase.Tab.ADDRESS_TO] = DefinedNames.GetDefaultAddressToTable(DataBase.Tab.ADDRESS_TO);
            this[DataBase.Tab.EDITABILI] = DefinedNames.GetDefaultEditableTable(DataBase.Tab.EDITABILI);
            this[DataBase.Tab.SALVADB] = DefinedNames.GetDefaultSaveTable(DataBase.Tab.SALVADB);
            this[DataBase.Tab.ANNOTA] = DefinedNames.GetDefaultToNoteTable(DataBase.Tab.ANNOTA);
            this[DataBase.Tab.CHECK] = DefinedNames.GetDefaultCheckTable(DataBase.Tab.CHECK);
            this[DataBase.Tab.SELECTION] = DefinedNames.GetDefaultSelectionTable(DataBase.Tab.SELECTION);
            this[DataBase.Tab.MODIFICA] = CreaTabellaModifica(DataBase.Tab.MODIFICA);
            this[DataBase.Tab.EXPORT_XML] = CreaTabellaExportXML(DataBase.Tab.EXPORT_XML);
        }
        private DataTable CreaTabellaModifica(string name)
        {
            DataTable dt = new DataTable(name)
            {
                Columns =
                {
                    {"SiglaEntita", typeof(string)},
                    {"SiglaInformazione", typeof(string)},
                    {"Data", typeof(string)},
                    {"Valore", typeof(string)},
                    {"AnnotaModifica", typeof(string)},
                    {"IdApplicazione", typeof(string)},
                    {"IdUtente", typeof(string)}
                }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["SiglaEntita"], dt.Columns["SiglaInformazione"], dt.Columns["Data"] };
            return dt;
        }
        private DataTable CreaTabellaExportXML(string name)
        {
            DataTable dt = new DataTable(name)
            {
                Columns =
                {
                    {"SiglaEntita", typeof(string)},
                    {"SiglaInformazione", typeof(string)},
                    {"Data", typeof(string)},
                    {"Valore", typeof(string)},
                    {"AnnotaModifica", typeof(string)},
                    {"IdApplicazione", typeof(string)},
                    {"IdUtente", typeof(string)}
                }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["SiglaEntita"], dt.Columns["SiglaInformazione"], dt.Columns["Data"] };
            return dt;        
        }
        
        #endregion

        /// <summary>
        /// Metodo richiamato da tutte le routine sottostanti che effettua la chiamata alla stored procedure sul server e aggiunge la tabella al DataSet locale. Restituisce true se l'operazione è andata a buon fine, lancia un'eccezione RepositoryUpdateException se fallisce.
        /// </summary>
        /// <param name="tableName">Nome della tabella da aggiornare.</param>
        /// <param name="spName">Nome della stored procedure da eseguire.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <returns>True se l'operazione è andata a buon fine.</returns>
        private void CaricaDati(string tableName, string spName, QryParams parameters)
        {
            DataTable dt = new DataTable();
            if (_isMultiApplication)
            {
                foreach (string id in _appIDs)
                {
                    parameters["@IdApplicazione"] = id;
                    dt.Merge(DataBase.Select(spName, parameters) ?? new DataTable());
                }
                if (dt.Columns.Count == 0)
                    dt = null;
            }
            else
            {
                dt = DataBase.Select(spName, parameters);
            }
            
            if (dt != null)
                this[tableName] = dt;
        }
        /// <summary>
        /// Carica la lista di tutte le applicazioni disponibili.
        /// </summary>
        private void CaricaApplicazioni()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@IdApplicazione", 0}
                };

            CaricaDati(DataBase.Tab.LISTA_APPLICAZIONI, DataBase.SP.APPLICAZIONE, parameters);
        }
        /// <summary>
        /// Carica i dati necessari alla creazione del menu ribbon.
        /// </summary>
        /// <returns></returns>
        private void CaricaApplicazioneRibbon()
        {
            CaricaDati(DataBase.Tab.APPLICAZIONE_RIBBON, DataBase.SP.APPLICAZIONE_RIBBON, new QryParams());
        }
        /// <summary>
        /// Carica le azioni.
        /// </summary>
        /// <returns></returns>
        private void CaricaAzioni()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@Operativa", Core.DataBase.ALL},
                    {"@Visibile", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.AZIONE, DataBase.SP.AZIONE, parameters);
        }
        /// <summary>
        /// Carica le categorie.
        /// </summary>
        /// <returns></returns>
        private void CaricaCategorie()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaCategoria", Core.DataBase.ALL},
                    {"@Operativa", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.CATEGORIA, DataBase.SP.CATEGORIA, parameters);
        }
        /// <summary>
        /// Carica la relazione azione categoria.
        /// </summary>
        /// <returns></returns>
        private void CaricaAzioneCategoria()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@SiglaCategoria", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.AZIONE_CATEGORIA, DataBase.SP.AZIONE_CATEGORIA, parameters);
        }
        /// <summary>
        /// Carica la relazione categoria entita.
        /// </summary>
        /// <returns></returns>
        private void CaricaCategoriaEntita()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaCategoria", Core.DataBase.ALL},
                    {"@SiglaEntita", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.CATEGORIA_ENTITA, DataBase.SP.CATEGORIA_ENTITA, parameters);
        }
        /// <summary>
        /// Carica la relazione entità azione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaAzione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaAzione", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.ENTITA_AZIONE, DataBase.SP.ENTITA_AZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità azione calcolo.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaAzioneCalcolo()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@SiglaCalcolo", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.ENTITA_AZIONE_CALCOLO, DataBase.SP.ENTITA_AZIONE_CALCOLO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità informazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaInformazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.ENTITA_INFORMAZIONE, DataBase.SP.ENTITA_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità azione informazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaAzioneInformazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.ENTITA_AZIONE_INFORMAZIONE, DataBase.SP.ENTITA_AZIONE_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica i calcoli.
        /// </summary>
        /// <returns></returns>
        private void CaricaCalcolo()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaCalcolo", Core.DataBase.ALL},
                    {"@IdTipologiaCalcolo", 0}
                };

            CaricaDati(DataBase.Tab.CALCOLO, DataBase.SP.CALCOLO, parameters);
        }
        /// <summary>
        /// Carica la relazione calcolo informazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaCalcoloInformazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaCalcolo", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.CALCOLO_INFORMAZIONE, DataBase.SP.CALCOLO_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità calcolo.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaCalcolo()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaCalcolo", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.ENTITA_CALCOLO, DataBase.SP.ENTITA_CALCOLO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità grafico.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaGrafico()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaGrafico", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.ENTITA_GRAFICO, DataBase.SP.ENTITA_GRAFICO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità grafico informazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaGraficoInformazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaGrafico", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.ENTITA_GRAFICO_INFORMAZIONE, DataBase.SP.ENTITA_GRAFICO_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità commitment.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaCommitment()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.ENTITA_COMMITMENT, DataBase.SP.ENTITA_COMMITMENT, parameters);
        }
        /// <summary>
        /// Carica la relazione entità rampa.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaRampa()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.ENTITA_RAMPA, DataBase.SP.ENTITA_RAMPA, parameters);
        }
        /// <summary>
        /// Carica la relazione entità assetto.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaAssetto()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.ENTITA_ASSETTO, DataBase.SP.ENTITA_ASSETTO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità proprietà.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaProprieta()
        {
            CaricaDati(DataBase.Tab.ENTITA_PROPRIETA, DataBase.SP.ENTITA_PROPRIETA, new QryParams());
        }
        /// <summary>
        /// Carica la relazione entità informazione formattazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaInformazioneFormattazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.Tab.ENTITA_INFORMAZIONE_FORMATTAZIONE, DataBase.SP.ENTITA_INFORMAZIONE_FORMATTAZIONE, parameters);
        }
        /// <summary>
        /// Carica la tipologia check.
        /// </summary>
        /// <returns></returns>
        private void CaricaTipologiaCheck()
        {
            CaricaDati(DataBase.Tab.TIPOLOGIA_CHECK, DataBase.SP.TIPOLOGIA_CHECK, new QryParams());
        }
        /// <summary>
        /// Carica la relazione entità parametro giornaliero.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaParametroD()
        {
            CaricaDati(DataBase.Tab.ENTITA_PARAMETRO_D, DataBase.SP.ENTITA_PARAMETRO_D, new QryParams());
        }
        /// <summary>
        /// Carica la relazione entità parametro orario.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaParametroH()
        {
            CaricaDati(DataBase.Tab.ENTITA_PARAMETRO_H, DataBase.SP.ENTITA_PARAMETRO_H, new QryParams());
        }

        #endregion

        /// <summary>
        /// Carica solo i parametri dell'applicazione contenuti nella tabella APPLICAZIONE.
        /// </summary>
        /// <param name="appID">L'ID dell'applicazione di cui ricaricare i parametri.</param>
        /// <returns>La tabella contenente tutti i dati trovati sull'applicazione.</returns>
        public DataRow CaricaApplicazione(object IdApplicazione)
        {
            CaricaApplicazioni();

            Applicazione = this[DataBase.Tab.LISTA_APPLICAZIONI].AsEnumerable()
                .Where(r => r["IdApplicazione"].Equals(IdApplicazione))
                .FirstOrDefault();

            return Applicazione;
        }

        #endregion

        public IEnumerator GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }

    class Win32Window : IWin32Window
    {
        public Win32Window(IntPtr handle) { Handle = handle; }
        public IntPtr Handle { get; private set; }
    }
}
