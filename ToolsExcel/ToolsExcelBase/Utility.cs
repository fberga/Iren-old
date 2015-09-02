using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Core;
using Iren.ToolsExcel.UserConfig;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
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
                APPLICAZIONE = "Applicazione",
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
        public static void InitNewDB(string dbName) 
        {
            _db = new Core.DataBase(dbName);
        }
        /// <summary>
        /// Inizializza il nuovo dataset Locale.
        /// </summary>
        public static void InitNewLocalDB()
        {
            _localDB = new DataSet(NAME);
        }
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
                string cartellaRemota = Esporta.PreparePath(path.Value);
                //path della cartella di emergenza
                string cartellaEmergenza = Esporta.PreparePath(path.Emergenza);
                //path della cartella di archivio in cui copiare i file in caso di esito positivo nel saltavaggio
                string cartellaArchivio = Esporta.PreparePath(path.Archivio);

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
        public static bool Insert(string storedProcedure, QryParams parameters)
        {
            if (OpenConnection())
            {
                bool o = _db.Insert(storedProcedure, parameters);
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

                    Excel.Worksheet ws = Workbook.Log;
                    if (ws.ListObjects.Count > 0)
                        ws.ListObjects[1].Range.EntireColumn.AutoFit();
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

    public class Repository : DataBase
    {
        #region Variabili

        private static bool _daAggiornare = false;
        private static bool _isMultiApplication = false;
        private static string[] _appIDs = null;

        #endregion

        #region Proprietà

        public static bool DaAggiornare { get { return _daAggiornare; } set { _daAggiornare = value; } }
        public static bool IsMultiApplicaion { get { return _isMultiApplication; } }

        #endregion

        #region Metodi

        public static void InitStrutturaNomi()
        {
            CreaTabellaNomi();
            CreaTabellaDate();
            CreaTabellaAddressFrom();
            CreaTabellaAddressTo();
            CreaTabellaModifica();
            CreaTabellaEditabili();
            CreaTabellaSalvaDB();
            CreaTabellaAnnotaModifica();
            CreaTabellaCheck();
            CreaTabellaSelection();
            _localDB.AcceptChanges();
        }

        /// <summary>
        /// Launcher per tutte le funzioni che aggiornano il repository in seguito alla richiesta di aggiornare la struttura.
        /// </summary>
        public static void Aggiorna(string[] appIDs)
        {
            _isMultiApplication = appIDs != null;
            _appIDs = appIDs;

            InitStrutturaNomi();
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
            
            _localDB.AcceptChanges();
        }
        #region Aggiorna Struttura Dati

        #region Init Struttura Nomi
        private static bool CreaTabellaNomi()
        {
            try
            {
                string name = Tab.NOMI_DEFINITI;
                ResetTable(name);
                DataTable dt = DefinedNames.GetDefaultNameTable(name);
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CreaTabellaDate()
        {
            try
            {
                string name = Tab.DATE_DEFINITE;
                ResetTable(name);
                DataTable dt = DefinedNames.GetDefaultDateTable(name);
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CreaTabellaAddressFrom()
        {
            try
            {
                string name = Tab.ADDRESS_FROM;
                ResetTable(name);
                DataTable dt = DefinedNames.GetDefaultAddressFromTable(name);
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CreaTabellaAddressTo()
        {
            try
            {
                string name = Tab.ADDRESS_TO;
                ResetTable(name);
                DataTable dt = DefinedNames.GetDefaultAddressToTable(name);
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CreaTabellaEditabili()
        {
            try
            {
                string name = Tab.EDITABILI;
                ResetTable(name);
                DataTable dt = DefinedNames.GetDefaultEditableTable(name);
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CreaTabellaSalvaDB()
        {
            try
            {
                string name = Tab.SALVADB;
                ResetTable(name);
                DataTable dt = DefinedNames.GetDefaultSaveTable(name);
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CreaTabellaAnnotaModifica()
        {
            try
            {
                string name = Tab.ANNOTA;
                ResetTable(name);
                DataTable dt = DefinedNames.GetDefaultToNoteTable(name);
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CreaTabellaModifica()
        {
            try
            {
                string name = Tab.MODIFICA;
                ResetTable(name);
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
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CreaTabellaCheck()
        {
            try
            {
                string name = Tab.CHECK;
                ResetTable(name);
                DataTable dt = DefinedNames.GetDefaultCheckTable(name);
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CreaTabellaSelection()
        {
            try
            {
                string name = Tab.SELECTION;
                ResetTable(name);
                DataTable dt = DefinedNames.GetDefaultSelectionTable(name);
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        #endregion

        /// <summary>
        /// Metodo richiamato da tutte le routine sottostanti che effettua la chiamata alla stored procedure sul server e aggiunge la tabella al DataSet locale. Restituisce true se l'operazione è andata a buon fine, lancia un'eccezione RepositoryUpdateException se fallisce.
        /// </summary>
        /// <param name="tableName">Nome della tabella da aggiornare.</param>
        /// <param name="spName">Nome della stored procedure da eseguire.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <returns>True se l'operazione è andata a buon fine.</returns>
        public static void CaricaDati(string tableName, string spName, QryParams parameters)
        {
            DataTable dt = new DataTable();
            if (_isMultiApplication)
            {
                foreach (string id in _appIDs)
                {
                    parameters["@IdApplicazione"] = id;
                    dt.Merge(Select(spName, parameters) ?? new DataTable());
                }
                if (dt.Columns.Count == 0)
                    dt = null;
            }
            else
            {
                dt = Select(spName, parameters);
            }
            if (dt != null)
            {
                ResetTable(tableName);
                dt.TableName = tableName;
                _localDB.Tables.Add(dt);
            }
        }
        /// <summary>
        /// Carica i dati necessari alla creazione del menu ribbon.
        /// </summary>
        /// <returns></returns>
        public static void CaricaApplicazioneRibbon()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@IdApplicazione", DB.IdApplicazione},
                    {"@IdUtente", DB.IdUtenteAttivo}
                };

            CaricaDati(Tab.APPLICAZIONE_RIBBON, SP.APPLICAZIONE_RIBBON, parameters);
        }
        /// <summary>
        /// Carica le azioni.
        /// </summary>
        /// <returns></returns>
        private static void CaricaAzioni()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@Operativa", Core.DataBase.ALL},
                    {"@Visibile", Core.DataBase.ALL}
                };

            CaricaDati(Tab.AZIONE, SP.AZIONE, parameters);
        }
        /// <summary>
        /// Carica le categorie.
        /// </summary>
        /// <returns></returns>
        private static void CaricaCategorie()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaCategoria", Core.DataBase.ALL},
                    {"@Operativa", Core.DataBase.ALL}
                };

            CaricaDati(Tab.CATEGORIA, SP.CATEGORIA, parameters);
        }
        /// <summary>
        /// Carica la relazione azione categoria.
        /// </summary>
        /// <returns></returns>
        private static void CaricaAzioneCategoria()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@SiglaCategoria", Core.DataBase.ALL}
                };

            CaricaDati(Tab.AZIONE_CATEGORIA, SP.AZIONE_CATEGORIA, parameters);
        }
        /// <summary>
        /// Carica la relazione categoria entita.
        /// </summary>
        /// <returns></returns>
        private static void CaricaCategoriaEntita()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaCategoria", Core.DataBase.ALL},
                    {"@SiglaEntita", Core.DataBase.ALL}
                };

            CaricaDati(Tab.CATEGORIA_ENTITA, SP.CATEGORIA_ENTITA, parameters);
        }
        /// <summary>
        /// Carica la relazione entità azione.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaAzione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaAzione", Core.DataBase.ALL}
                };

            CaricaDati(Tab.ENTITA_AZIONE, SP.ENTITA_AZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità azione calcolo.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaAzioneCalcolo()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@SiglaCalcolo", Core.DataBase.ALL}
                };

            CaricaDati(Tab.ENTITA_AZIONE_CALCOLO, SP.ENTITA_AZIONE_CALCOLO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità informazione.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaInformazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(Tab.ENTITA_INFORMAZIONE, SP.ENTITA_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità azione informazione.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaAzioneInformazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(Tab.ENTITA_AZIONE_INFORMAZIONE, SP.ENTITA_AZIONE_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica i calcoli.
        /// </summary>
        /// <returns></returns>
        private static void CaricaCalcolo()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaCalcolo", Core.DataBase.ALL},
                    {"@IdTipologiaCalcolo", 0}
                };

            CaricaDati(Tab.CALCOLO, SP.CALCOLO, parameters);
        }
        /// <summary>
        /// Carica la relazione calcolo informazione.
        /// </summary>
        /// <returns></returns>
        private static void CaricaCalcoloInformazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaCalcolo", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(Tab.CALCOLO_INFORMAZIONE, SP.CALCOLO_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità calcolo.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaCalcolo()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaCalcolo", Core.DataBase.ALL}
                };

            CaricaDati(Tab.ENTITA_CALCOLO, SP.ENTITA_CALCOLO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità grafico.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaGrafico()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaGrafico", Core.DataBase.ALL}
                };

            CaricaDati(Tab.ENTITA_GRAFICO, SP.ENTITA_GRAFICO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità grafico informazione.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaGraficoInformazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaGrafico", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(Tab.ENTITA_GRAFICO_INFORMAZIONE, SP.ENTITA_GRAFICO_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità commitment.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaCommitment()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL}
                };

            CaricaDati(Tab.ENTITA_COMMITMENT, SP.ENTITA_COMMITMENT, parameters);
        }
        /// <summary>
        /// Carica la relazione entità rampa.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaRampa()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL}
                };

            CaricaDati(Tab.ENTITA_RAMPA, SP.ENTITA_RAMPA, parameters);
        }
        /// <summary>
        /// Carica la relazione entità assetto.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaAssetto()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL}
                };

            CaricaDati(Tab.ENTITA_ASSETTO, SP.ENTITA_ASSETTO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità proprietà.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaProprieta()
        {
            CaricaDati(Tab.ENTITA_PROPRIETA, SP.ENTITA_PROPRIETA, new QryParams());
        }
        /// <summary>
        /// Carica la relazione entità informazione formattazione.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaInformazioneFormattazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(Tab.ENTITA_INFORMAZIONE_FORMATTAZIONE, SP.ENTITA_INFORMAZIONE_FORMATTAZIONE, parameters);
        }
        /// <summary>
        /// Carica la tipologia check.
        /// </summary>
        /// <returns></returns>
        private static void CaricaTipologiaCheck()
        {
            CaricaDati(Tab.TIPOLOGIA_CHECK, SP.TIPOLOGIA_CHECK, new QryParams());
        }
        /// <summary>
        /// Carica la relazione entità parametro giornaliero.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaParametroD()
        {
            CaricaDati(Tab.ENTITA_PARAMETRO_D, SP.ENTITA_PARAMETRO_D, new QryParams());
        }
        /// <summary>
        /// Carica la relazione entità parametro orario.
        /// </summary>
        /// <returns></returns>
        private static void CaricaEntitaParametroH()
        {
            CaricaDati(Tab.ENTITA_PARAMETRO_H, SP.ENTITA_PARAMETRO_H, new QryParams());
        }

        #endregion

        /// <summary>
        /// Carica solo i parametri dell'applicazione contenuti nella tabella APPLICAZIONE.
        /// </summary>
        /// <param name="appID">L'ID dell'applicazione di cui ricaricare i parametri.</param>
        /// <returns>La tabella contenente tutti i dati trovati sull'applicazione.</returns>
        public static DataTable CaricaApplicazione(object appID)
        {
            string name = DataBase.Tab.APPLICAZIONE;
            DataBase.ResetTable(name);
            QryParams parameters = new QryParams() 
            {
                {"@IdApplicazione", appID},

            };
            DataTable dt = DataBase.Select(DataBase.SP.APPLICAZIONE, parameters) ?? new DataTable();
            
            dt.TableName = name;
            return dt;
        }

        #endregion
    }

    public class Workbook 
    {
        #region Variabili

        /// <summary>
        /// Il workbook.
        /// </summary>
        protected static Microsoft.Office.Tools.Excel.Workbook _wb;
        /// <summary>
        /// La versione dell'applicazione.
        /// </summary>
        private static System.Version _wbVersion;
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
        public static Microsoft.Office.Tools.Excel.Workbook WB { get { return _wb; } }
        /// <summary>
        /// Scorciatoia per accedere all'oggetto Excel del foglio Main (sempre presente in tutti i fogli).
        /// </summary>
        public static Excel.Worksheet Main { get { return _wb.Sheets["Main"]; } }
        /// <summary>
        /// Scorciatoia per accedere all'oggetto Excel del foglio di Log (sempre presente in tutti i fogli).
        /// </summary>
        public static Excel.Worksheet Log { get { return _wb.Sheets["Log"]; } }
        /// <summary>
        /// Scorciatoia per accedere al foglio attivo.
        /// </summary>
        public static Excel.Worksheet ActiveSheet { get { return _wb.ActiveSheet; } }
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
        public static Excel.Sheets Sheets { get { return _wb.Sheets; } }
        /// <summary>
        /// Lista dei folgi MSDx utile solo in Invio Programmi.
        /// </summary>
        public static IList<Excel.Worksheet> MSDSheets { get { return _wb.Sheets.Cast<Excel.Worksheet>().Where(ws => ws.Name.StartsWith("MSD")).ToList(); } }
        /// <summary>
        /// La versione dell'applicazione.
        /// </summary>
        public static System.Version WorkbookVersion { get { return _wbVersion; } }
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

        #endregion

        #region Metodi

        /// <summary>
        /// Carica dal DB i dati riguardanti le proprietà dell'applicazione che si trovano nella tabella APPLICAZIONE. Assegna alle variabili globali di applicazione i valori.
        /// </summary>
        public static void AggiornaParametriApplicazione()
        {
            DataTable dt = Repository.CaricaApplicazione(Workbook.AppSettings("AppID"));
            if (dt.Rows.Count == 0)
                throw new ApplicationNotFoundException("L'appID inserito non ha restituito risultati.");

            Simboli.nomeApplicazione = dt.Rows[0]["DesApplicazione"].ToString();
            Struct.intervalloGiorni = (dt.Rows[0]["IntervalloGiorniEntita"] is DBNull ? 0 : (int)dt.Rows[0]["IntervalloGiorniEntita"]);
            Struct.tipoVisualizzazione = dt.Rows[0]["TipoVisualizzazione"] is DBNull ? "O" : dt.Rows[0]["TipoVisualizzazione"].ToString();
            Struct.visualizzaRiepilogo = dt.Rows[0]["VisRiepilogo"] is DBNull ? true : dt.Rows[0]["VisRiepilogo"].Equals("1");

            Struct.cell.width.empty = double.Parse(dt.Rows[0]["ColVuotaWidth"].ToString());
            Struct.cell.width.dato = double.Parse(dt.Rows[0]["ColDatoWidth"].ToString());
            Struct.cell.width.entita = double.Parse(dt.Rows[0]["ColEntitaWidth"].ToString());
            Struct.cell.width.informazione = double.Parse(dt.Rows[0]["ColInformazioneWidth"].ToString());
            Struct.cell.width.unitaMisura = double.Parse(dt.Rows[0]["ColUMWidth"].ToString());
            Struct.cell.width.parametro = double.Parse(dt.Rows[0]["ColParametroWidth"].ToString());
            Struct.cell.width.jolly1 = double.Parse(dt.Rows[0]["ColJolly1Width"].ToString());
            Struct.cell.height.normal = double.Parse(dt.Rows[0]["RowHeight"].ToString());
            Struct.cell.height.empty = double.Parse(dt.Rows[0]["RowVuotaHeight"].ToString());

            DataBase.ResetTable(DataBase.Tab.APPLICAZIONE);
            DataBase.LocalDB.Tables.Add(dt);
        }
        /// <summary>
        /// Imposta il mercato attivo in base all'orario. Se necessario cambia anche la data attiva e imposta il foglio come da aggiornare.
        /// </summary>
        /// <param name="appID">L'ID applicazione che identifica anche in quale mercato il foglio è impostato.</param>
        /// <param name="dataAttiva">La data attiva da modificare all'occorrenza.</param>
        /// <returns>Restituisce true se il foglio è da aggiornare, false altrimenti.</returns>
        private static bool SetMercato(ref string appID, ref DateTime dataAttiva)
        {
            string appIDold = appID;
            DateTime dataAttivaOld = dataAttiva;

            //configuro la data attiva
            int ora = DateTime.Now.Hour;
            if (ora > 17)
                dataAttiva = DateTime.Today.AddDays(1);
            else if (ora >= 7 && ora <= 17)
                dataAttiva = DateTime.Today;

            //configuro il mercato attivo
            string[] mercatiDisp = Workbook.AppSettings("Mercati").Split('|');
            string[] appIDs = Workbook.AppSettings("AppIDMSD").Split('|');
            for(int i = 0; i < mercatiDisp.Length; i++) 
            {
                string[] ore = Workbook.AppSettings("Ore" + mercatiDisp[i]).Split('|');
                if (ore.Contains(ora.ToString()))
                {
                    appID = appIDs[i];
                    break;
                }
            }

            Simboli.AppID = appID;

            if(appID != appIDold || dataAttivaOld != dataAttiva)
            {
                Workbook.ChangeAppSettings("DataAttiva", dataAttiva.ToString("yyyyMMdd"));
                Simboli.AppID = appID;

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
        private static bool AggiornaData(string appID, ref DateTime dataAttiva)
        {            
            DateTime dataAttivaOld = dataAttiva;

            if (appID == "12")
            {
                //configuro la data attiva
                int ora = DateTime.Now.Hour;
                if (ora <= 15)
                    dataAttiva = DateTime.Today.AddDays(1);
                else
                    dataAttiva = DateTime.Today.AddDays(2);
            }
            else
            {
                dataAttiva = DateTime.Today.AddDays(-1);
            }
            

            if (dataAttivaOld != dataAttiva)
            {
                Workbook.ChangeAppSettings("DataAttiva", dataAttiva.ToString("yyyyMMdd"));
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

                    if (DataBase.OpenConnection())
                    {
                        Dictionary<Core.DataBase.NomiDB, ConnectionState> stato = DataBase.StatoDB;
                        Simboli.SQLServerOnline = stato[Core.DataBase.NomiDB.SQLSERVER] == ConnectionState.Open;
                        Simboli.ImpiantiOnline = stato[Core.DataBase.NomiDB.IMP] == ConnectionState.Open;
                        Simboli.ElsagOnline = stato[Core.DataBase.NomiDB.ELSAG] == ConnectionState.Open;

                        DataBase.CloseConnection();
                    }
                    else
                    {
                        Simboli.SQLServerOnline = false;
                        Simboli.ImpiantiOnline = false;
                        Simboli.ElsagOnline = false;
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
            try { _wb.CustomXMLParts[WB.Name].Delete(); }
            catch { }
            part = _wb.CustomXMLParts.Add();
            //carico nella nuova custom part il contenuto.
            part.LoadXML(root.ToString(SaveOptions.DisableFormatting));
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
            if(dtLog != null)
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
        private static int InitUser()
        {
            DataTable dtUtente = DataBase.Select(DataBase.SP.UTENTE, new QryParams() { { "@CodUtenteWindows", Environment.UserName } });
            if (dtUtente != null)
            {
                dtUtente.TableName = DataBase.Tab.UTENTE;

                if (dtUtente.Rows.Count == 0)
                {
                    DataRow r = dtUtente.NewRow();
                    r["IdUtente"] = 0;
                    r["Nome"] = "NON CONFIGURATO";
                    dtUtente.Rows.Add(r);
                }
                    
                DataBase.ResetTable(DataBase.Tab.UTENTE);
                DataBase.LocalDB.Tables.Add(dtUtente);

                return int.Parse("" + dtUtente.Rows[0]["IdUtente"]);
            }

            System.Windows.Forms.MessageBox.Show("Errore durante l'inizializzazione dell'utente.", Simboli.nomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            
            return -1;            
        }
        private static bool Init(string dbName, string appID, DateTime dataAttiva)
        {
            //CryptHelper.CryptSection("connectionStrings", "appSettings");

            //controllo le aree di rete (se presenti)
            var usrConfig = GetUsrConfiguration();
            Dictionary<string, string> pathNonDisponibili = new Dictionary<string, string>();
            foreach (UserConfigElement ele in usrConfig.Items)
            {
                if (ele.ToCheckPath == "true")
                {
                    string pathStr = Esporta.PreparePath(ele.Value);

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


            DataBase.InitNewDB(dbName);
            DataBase.DB.PropertyChanged += _db_StatoDBChanged;
            DataBase.InitNewLocalDB();

            bool localDBNotPresent = false;
            try
            {
                Office.CustomXMLPart xmlPart = _wb.CustomXMLParts[WB.Name];
                StringReader sr = new StringReader(xmlPart.XML);
                DataBase.LocalDB.ReadXml(sr);
            }
            catch
            {
                localDBNotPresent = true;
                DataBase.LocalDB.Namespace = WB.Name;
            }

            bool toUpdate = false;

            //per Invio Programmi
            if (Workbook.AppSettings("Mercati") != null)
                toUpdate = SetMercato(ref appID, ref dataAttiva);

            //per Previsione Carico Termico & Validazione Teleriscaldamento
            if (appID == "11" || appID == "12")
                toUpdate = AggiornaData(appID, ref dataAttiva);

            if (DataBase.OpenConnection())
            {
                Workbook.AggiornaParametriApplicazione();

                int usr = InitUser();
                DataBase.DB.SetParameters(dataAttiva.ToString("yyyyMMdd"), usr, int.Parse(appID));

                DataView applicazione = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE].DefaultView;

                Simboli.rgbSfondo = Workbook.GetRGBFromString(applicazione[0]["BackColorApp"].ToString());
                Simboli.rgbTitolo = Workbook.GetRGBFromString(applicazione[0]["BackColorFrameApp"].ToString());
                Simboli.rgbLinee = Workbook.GetRGBFromString(applicazione[0]["BorderColorApp"].ToString());

                InitLog();

                Repository.DaAggiornare = toUpdate;

                return false;
            }
            else //Emergenza
            {
                if (localDBNotPresent)
                {
                    System.Windows.Forms.MessageBox.Show("Il foglio non è inizializzato e non c'è connessione ad DB... Impossibile procedere! L'applicazione verrà chiusa.", "ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                    _wb.Close();
                    return false;
                }

                DataBase.DB.SetParameters(dataAttiva.ToString("yyyyMMdd"), 0, 0);
                DataView applicazione = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE].DefaultView;
                Simboli.nomeApplicazione = applicazione[0]["DesApplicazione"].ToString();
                Struct.intervalloGiorni = applicazione[0]["IntervalloGiorniEntita"] is DBNull ? 0 : int.Parse(applicazione[0]["IntervalloGiorniEntita"].ToString());
                Struct.visualizzaRiepilogo = applicazione[0]["VisRiepilogo"] is DBNull ? true : applicazione[0]["VisRiepilogo"].Equals("1");

                return true;
            }
        }
        public static void StartUp(Microsoft.Office.Tools.Excel.Workbook wb, System.Version wbVersion)
        {
            _wb = wb;
            _wbVersion = wbVersion;

            Window = new Win32Window(new IntPtr(Workbook.Application.Hwnd));


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
                if (ws.Name != "Main")
                    Application.ActiveWindow.ScrollColumn = 1;
            }

            Main.Select();
            Application.WindowState = Excel.XlWindowState.xlMaximized;

            Simboli.pwd = AppSettings("pwd");

            bool wasProtected = Sheet.Protected;
            if (wasProtected)
                Sheet.Protected = false;

            Workbook.ScreenUpdating = false;

            DateTime dataAttiva = DateTime.ParseExact(AppSettings("DataAttiva"), "yyyyMMdd", CultureInfo.InvariantCulture);
            bool emergenza = Init(AppSettings("DB"), AppSettings("AppID"), dataAttiva);

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
            _wb.Save();
        }
        public static void Close()
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
            Save();

            Application.ScreenUpdating = true;

            //Window.ReleaseHandle();
        }
        #endregion

        #endregion
    }

    class Win32Window : IWin32Window
    {
        public Win32Window(IntPtr handle) { Handle = handle; }
        public IntPtr Handle { get; private set; }
    }
}
