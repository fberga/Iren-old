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
                TIPOLOGIA_RAMPA = "spTipologiaRampa",
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
                TIPOLOGIA_RAMPA = "TipologiaRampa",
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

        #region Metodi

        public static void InitNewDB(string dbName) {
            _db = new Core.DataBase(dbName);
        }
        public static void InitNewLocalDB()
        {
            _localDB = new DataSet(NAME);
        }
        public static void ResetTable(string name) 
        {
            if (_localDB.Tables.Contains(name))
            {
                _localDB.Tables.Remove(name);
            }
        }
        public static void ChangeAppID(string appID)
        {
            Utility.DataBase.ChangeAppSettings("AppID", appID);
            _db.ChangeAppID(int.Parse(appID));
        }
        public static void ChangeDate(DateTime dataAttiva) 
        {
            _db.ChangeDate(dataAttiva.ToString("yyyyMMdd"));
        }
        public static void ChangeAppSettings(string key, string value) 
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings[key].Value = value;
            config.Save(ConfigurationSaveMode.Minimal);
            ConfigurationManager.RefreshSection("applicationSettings");
        }
        public static void SwitchEnvironment(string ambiente) 
        {            
            ChangeAppSettings("DB", ambiente);
            ConfigurationManager.RefreshSection("applicationSettings");

            int idA = _db.IdApplicazione;
            int idU = _db.IdUtenteAttivo;
            string data = _db.DataAttiva.ToString("yyyyMMdd");

            _db = new Core.DataBase(ambiente);
            _db.SetParameters(data, idU, idA);
        }
        public static void SalvaModificheDB() 
        {
            DataTable modifiche = LocalDB.Tables[Tab.MODIFICA];
            if (modifiche != null)
            {
                DataTable dt = modifiche.Copy();
                dt.TableName = modifiche.TableName;
                dt.Namespace = "";

                if (dt.Rows.Count == 0)
                    return;

                bool onLine = DB.OpenConnection();

                var path = Workbook.GetUsrConfigElement("pathExportModifiche");

                string cartellaRemota = ExportPath.PreparePath(path.Value);
                string cartellaEmergenza = ExportPath.PreparePath(path.Emergenza);
                string cartellaArchivio = ExportPath.PreparePath(path.Archivio);

                string fileName = "";
                if (onLine && Directory.Exists(cartellaRemota))
                {
                    string[] fileEmergenza = Directory.GetFiles(cartellaEmergenza);
                    bool executed = false;
                    if (fileEmergenza.Length > 0)
                    {
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
                        }
                    }

                    fileName = Path.Combine(cartellaRemota, Simboli.nomeApplicazione.Replace(" ", "").ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml");
                    dt.WriteXml(fileName);

                    executed = DataBase.DB.Insert(SP.INSERT_APPLICAZIONE_INFORMAZIONE_XML, new QryParams() { { "@NomeFile", fileName.Split('\\').Last() } });
                    if (executed)
                    {
                        if (!Directory.Exists(cartellaArchivio))
                            Directory.CreateDirectory(cartellaArchivio);

                        File.Move(fileName, Path.Combine(cartellaArchivio, fileName.Split('\\').Last()));
                    }
                }
                else
                {
                    fileName = Path.Combine(cartellaEmergenza, Simboli.nomeApplicazione.Replace(" ", "").ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml");
                    try
                    {
                        dt.WriteXml(fileName, XmlWriteMode.IgnoreSchema);
                    }
                    catch (DirectoryNotFoundException)
                    {
                        Directory.CreateDirectory(cartellaEmergenza);
                        dt.WriteXml(fileName, XmlWriteMode.IgnoreSchema);
                    }
                }

                modifiche.Clear();
            }
        }
        //public static object GetMessaggioCheck(object id) 
        //{
        //    DataView tipologiaCheck = _localDB.Tables[Tab.TIPOLOGIA_CHECK].DefaultView;
        //    tipologiaCheck.RowFilter = "IdTipologiaCheck = " + id;

        //    if (tipologiaCheck.Count > 0)
        //        return tipologiaCheck[0]["Messaggio"];

        //    return null;
        //}
        public static void InsertApplicazioneRiepilogo(object siglaEntita, object siglaAzione, DateTime? dataRif = null, bool presente = true) 
        {
            dataRif = dataRif ?? DataAttiva;
            try
            {
                if (OpenConnection())
                {
                    QryParams parameters = new QryParams() {
                    {"@SiglaEntita", siglaEntita},
                    {"@SiglaAzione", siglaAzione},
                    {"@Data", dataRif.Value.ToString("yyyyMMdd")},
                    {"@Presente", presente ? "1" : "0"}
                };
                    _db.Insert(DataBase.SP.INSERT_APPLICAZIONE_RIEPILOGO, parameters);
                }
            }
            catch (Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "InsertApplicazioneRiepilogo [" + dataRif ?? DataAttiva + ", " + siglaEntita + ", " + siglaAzione + "]: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        public static void ConvertiParametriInformazioni() 
        {
            Select(SP.APPLICAZIONE_INIT);
        }
        public static bool OpenConnection()
        {
            if (!Simboli.EmergenzaForzata)
                return _db.OpenConnection();

            return false;
        }
        public static bool CloseConnection()
        {
            if (!Simboli.EmergenzaForzata)
                return _db.CloseConnection();

            return false;
        }
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
        public void RefreshLog()
        {
            if (OpenConnection())
            {
                DataTable dt = Select(SP.APPLICAZIONE_LOG);
                dt.TableName = Tab.LOG;

                bool sameSchema = dt.Columns.Count == _localDB.Tables[Tab.LOG].Columns.Count;

                for(int i = 0; i < dt.Columns.Count && sameSchema; i++)
                    if(_localDB.Tables[Tab.LOG].Columns[i].ColumnName != dt.Columns[i].ColumnName)
                        sameSchema = false;

                _localDB.Tables[Tab.LOG].Clear();

                //if (!sameSchema)
                //{
                //    while (_localDB.Tables[Tab.LOG].Columns.Count > 0)
                //        _localDB.Tables[Tab.LOG].Columns.RemoveAt(0);
                //    foreach (DataColumn col in dt.Columns)
                //        _localDB.Tables[Tab.LOG].Columns.Add(new DataColumn() { ColumnName = col.ColumnName, DataType = col.DataType });
                //}

                _localDB.Tables[Tab.LOG].Merge(dt);

                Excel.Worksheet ws = Workbook.Log;
                if (ws.ListObjects.Count > 0)
                    ws.ListObjects[1].Range.EntireColumn.AutoFit();
            }
        }

        public static DataTable Select(string storedProcedure, QryParams parameters, int timeout = 300)
        {
            if (OpenConnection())
            {
                DataTable dt = _db.Select(storedProcedure, parameters, timeout);
                CloseConnection();
                
                return dt;
            }

            return new DataTable();
        }
        public static DataTable Select(string storedProcedure, String parameters, int timeout = 300)
        {
            if (OpenConnection())
            {
                DataTable dt = _db.Select(storedProcedure, parameters, timeout);
                CloseConnection();

                return dt;
            }

            return new DataTable();
        }
        public static DataTable Select(string storedProcedure, int timeout = 300)
        {
            if (OpenConnection())
            {
                DataTable dt = _db.Select(storedProcedure, timeout);
                CloseConnection();

                return dt;
            }

            return new DataTable();
        }

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
    }

    public class ExportPath 
    {
        #region Metodi

        public static string PreparePath(string path, string codRup = "") 
        {
            Regex options = new Regex(@"\[\w+\]");
            path = options.Replace(path, match =>
            {
                string opt = match.Value.Replace("[", "").Replace("]", "");
                string o = "";
                switch (opt.ToLowerInvariant())
                {
                    case "appname":
                        o = Simboli.nomeApplicazione.Replace(" ", "").ToUpperInvariant();
                        break;
                    case "msd":
                        o = Simboli.Mercato;
                        break;
                    case "codrup":
                        o = codRup;
                        break;
                    //aggiungere qui tutti i formati data da considerare nella forma
                    //case "formato data":
                    case "yyyymmdd":
                        o = DataBase.DataAttiva.ToString(opt);
                        break;
                }

                return o;
            });

            return path;
        }

        #endregion
    }

    public class Date 
    {
        #region Proprietà

        public static string SuffissoDATA1
        {
            get { return GetSuffissoData(DataBase.DataAttiva); }
        }

        #endregion

        #region Metodi
        public static int GetOreIntervallo(DateTime fine)
        {
            return GetOreIntervallo(DataBase.DataAttiva, fine);
        }
        public static int GetOreIntervallo(DateTime inizio, DateTime fine)
        {
            return (int)(fine.AddDays(1).ToUniversalTime() - inizio.ToUniversalTime()).TotalHours;
        }
        public static int GetOreGiorno(DateTime giorno)
        {
            DateTime giornoSucc = giorno.AddDays(1);
            return (int)(giornoSucc.ToUniversalTime() - giorno.ToUniversalTime()).TotalHours;
        }
        public static int GetOreGiorno(string suffissoData)
        {
            return GetOreGiorno(GetDataFromSuffisso(suffissoData));
        }
        public static string GetSuffissoData(DateTime giorno)
        {
            return GetSuffissoData(Utility.DataBase.DataAttiva, giorno);
        }
        public static string GetSuffissoData(string giorno)
        {
            return GetSuffissoData(Utility.DataBase.DataAttiva, giorno);
        }
        public static string GetSuffissoData(DateTime inizio, DateTime giorno)
        {
            if (inizio > giorno)
            {
                return "DATA0";
            }
            TimeSpan dayDiff = giorno - inizio;
            return "DATA" + (dayDiff.Days + 1);
        }
        public static string GetSuffissoData(DateTime inizio, object giorno)
        {
            DateTime day = DateTime.ParseExact(giorno.ToString().Substring(0, 8), "yyyyMMdd", CultureInfo.InvariantCulture);
            return GetSuffissoData(inizio, day);
        }
        public static string GetSuffissoOra(int ora)
        {
            return "H" + ora;
        }
        public static string GetSuffissoOra(object dataOra)
        {
            string dtO = dataOra.ToString();
            if (dtO.Length != 10)
                return "";

            return GetSuffissoOra(int.Parse(dtO.Substring(dtO.Length - 2, 2)));
        }
        public static string GetDataFromSuffisso(string data, string ora)
        {
            DateTime outDate = GetDataFromSuffisso(data);
            ora = ora == "" ? "0" : ora;
            int outOra = int.Parse(Regex.Match(ora, @"\d+").Value);

            return outDate.ToString("yyyyMMdd") + (outOra != 0 ? outOra.ToString("D2") : "");
        }
        public static DateTime GetDataFromSuffisso(string data)
        {
            int giorno = int.Parse(Regex.Match(data.ToString(), @"\d+").Value);
            return DataBase.DB.DataAttiva.AddDays(giorno - 1);
        }
        public static int GetOraFromDataOra(string dataOra)
        {
            string dtO = dataOra.ToString();
            if (dtO.Length != 10)
                return -1;

            return int.Parse(dtO.Substring(dtO.Length - 2, 2));
        }
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

        #endregion

        #region Proprietà

        public static bool DaAggiornare { get { return _daAggiornare; } set { _daAggiornare = value; } }

        #endregion

        #region Metodi

        public static void Aggiorna(params string[] sheets)
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
            CaricaTipologiaCheck();
            CaricaTipologiaRampa();
            CaricaEntitaParametroD();
            CaricaEntitaParametroH();
            _localDB.AcceptChanges();
        }
        #region Aggiorna Struttura Dati

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
        private static bool CaricaPathApplicativi()
        {
            try
            {
                string name = Tab.AZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@Operativa", Core.DataBase.ALL},
                    {"@Visibile", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.AZIONE, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public static bool CaricaApplicazioneRibbon()
        {
            try
            {
                string name = Tab.APPLICAZIONE_RIBBON;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@IdApplicazione", DB.IdApplicazione},
                    {"@IdUtente", DB.IdUtenteAttivo}
                };
                DataTable dt = Select(SP.APPLICAZIONE_RIBBON, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaAzioni()
        {
            try
            {
                string name = Tab.AZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@Operativa", Core.DataBase.ALL},
                    {"@Visibile", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.AZIONE, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaCategorie()
        {
            try
            {
                string name = Tab.CATEGORIA;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaCategoria", Core.DataBase.ALL},
                    {"@Operativa", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.CATEGORIA, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaAzioneCategoria()
        {
            try
            {
                string name = Tab.AZIONE_CATEGORIA;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@SiglaCategoria", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.AZIONE_CATEGORIA, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaCategoriaEntita()
        {
            try
            {
                string name = Tab.CATEGORIA_ENTITA;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaCategoria", Core.DataBase.ALL},
                    {"@SiglaEntita", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.CATEGORIA_ENTITA, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaAzione()
        {
            try
            {
                string name = Tab.ENTITA_AZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaAzione", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.ENTITA_AZIONE, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaAzioneCalcolo()
        {
            try
            {
                string name = Tab.ENTITA_AZIONE_CALCOLO;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@SiglaCalcolo", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.ENTITA_AZIONE_CALCOLO, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaInformazione()
        {
            try
            {
                string name = Tab.ENTITA_INFORMAZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.ENTITA_INFORMAZIONE, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaAzioneInformazione()
        {
            try
            {
                string name = Tab.ENTITA_AZIONE_INFORMAZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.ENTITA_AZIONE_INFORMAZIONE, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaCalcolo()
        {
            try
            {
                string name = Tab.CALCOLO;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaCalcolo", Core.DataBase.ALL},
                    {"@IdTipologiaCalcolo", 0}
                };
                DataTable dt = Select(SP.CALCOLO, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaCalcoloInformazione()
        {
            try
            {
                string name = Tab.CALCOLO_INFORMAZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaCalcolo", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.CALCOLO_INFORMAZIONE, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaCalcolo()
        {
            try
            {
                string name = Tab.ENTITA_CALCOLO;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaCalcolo", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.ENTITA_CALCOLO, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaGrafico()
        {
            try
            {
                string name = Tab.ENTITA_GRAFICO;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaGrafico", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.ENTITA_GRAFICO, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaGraficoInformazione()
        {
            try
            {
                string name = Tab.ENTITA_GRAFICO_INFORMAZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaGrafico", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.ENTITA_GRAFICO_INFORMAZIONE, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaCommitment()
        {
            try
            {
                string name = Tab.ENTITA_COMMITMENT;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.ENTITA_COMMITMENT, parameters);
                dt.TableName = name;
                dt.PrimaryKey = new DataColumn[] { dt.Columns["SiglaEntita"], dt.Columns["SiglaCommitment"]};
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaRampa()
        {
            try
            {
                string name = Tab.ENTITA_RAMPA;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.ENTITA_RAMPA, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaAssetto()
        {
            try
            {
                string name = Tab.ENTITA_ASSETTO;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.ENTITA_ASSETTO, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaProprieta()
        {
            try
            {
                string name = Tab.ENTITA_PROPRIETA;
                ResetTable(name);
                QryParams parameters = new QryParams();
                DataTable dt = Select(SP.ENTITA_PROPRIETA, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaInformazioneFormattazione()
        {
            try
            {
                string name = Tab.ENTITA_INFORMAZIONE_FORMATTAZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };
                DataTable dt = Select(SP.ENTITA_INFORMAZIONE_FORMATTAZIONE, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaTipologiaCheck()
        {
            try
            {
                string name = Tab.TIPOLOGIA_CHECK;
                ResetTable(name);
                QryParams parameters = new QryParams();
                DataTable dt = Select(SP.TIPOLOGIA_CHECK, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaTipologiaRampa()
        {
            try
            {
                string name = Tab.TIPOLOGIA_RAMPA;
                ResetTable(name);
                QryParams parameters = new QryParams();
                DataTable dt = Select(SP.TIPOLOGIA_RAMPA, parameters);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaParametroD()
        {
            try
            {
                string name = Tab.ENTITA_PARAMETRO_D;
                ResetTable(name);
                DataTable dt = Select(SP.ENTITA_PARAMETRO_D);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private static bool CaricaEntitaParametroH()
        {
            try
            {
                string name = Tab.ENTITA_PARAMETRO_H;
                ResetTable(name);
                DataTable dt = Select(SP.ENTITA_PARAMETRO_H);
                dt.TableName = name;
                _localDB.Tables.Add(dt);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion


        public static DataTable CaricaApplicazione(object idApplicazione)
        {
            string name = DataBase.Tab.APPLICAZIONE;
            DataBase.ResetTable(name);
            QryParams parameters = new QryParams() 
            {
                {"@IdApplicazione", idApplicazione},

            };
            DataTable dt = DataBase.Select(DataBase.SP.APPLICAZIONE, parameters);
            dt.TableName = name;
            return dt;
        }

        #endregion
    }

    public class Workbook 
    {
        #region Variabili

        protected static Microsoft.Office.Tools.Excel.Workbook _wb;
        private static System.Version _wbVersion;
        public static bool fromErrorPane = false;

        #endregion

        #region Proprietà

        public static Microsoft.Office.Tools.Excel.Workbook WB { get { return _wb; } }
        public static Excel.Worksheet Main { get { return _wb.Sheets["Main"]; } }
        public static Excel.Worksheet Log { get { return _wb.Sheets["Log"]; } }
        public static Excel.Worksheet ActiveSheet { get { return _wb.ActiveSheet; } }
        public static Excel.Application Application { get { return _wb.Application; } }
        public static IList<Excel.Worksheet> CategorySheets { get { return _wb.Sheets.Cast<Excel.Worksheet>().Where(ws => ws.Name != "Log" && ws.Name != "Main" && !ws.Name.StartsWith("MSD")).ToList(); } }
        public static Excel.Sheets Sheets { get { return _wb.Sheets; } }
        public static IList<Excel.Worksheet> MSDSheets { get { return _wb.Sheets.Cast<Excel.Worksheet>().Where(ws => ws.Name.StartsWith("MSD")).ToList(); } }
        public static System.Version WorkbookVersion { get { return _wbVersion; } }
        public static System.Version CoreVersion { get { return DataBase.DB.GetCurrentV(); } }
        public static System.Version BaseVersion { get { return Assembly.GetExecutingAssembly().GetName().Version; } }
        public static bool ScreenUpdating { get { return Application.ScreenUpdating; } set { Application.ScreenUpdating = value; } }

        #endregion

        #region Metodi

        
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
            string[] mercatiDisp = ConfigurationManager.AppSettings["Mercati"].Split('|');
            string[] appIDs = ConfigurationManager.AppSettings["AppIDMSD"].Split('|');
            for(int i = 0; i < mercatiDisp.Length; i++) 
            {
                string[] ore = ConfigurationManager.AppSettings["Ore" + mercatiDisp[i]].Split('|');
                if (ore.Contains(ora.ToString()))
                {
                    appID = appIDs[i];
                    break;
                }
            }

            Simboli.AppID = appID;

            if(appID != appIDold || dataAttivaOld != dataAttiva)
            {
                DataBase.ChangeAppSettings("DataAttiva", dataAttiva.ToString("yyyyMMdd"));
                Simboli.AppID = appID;

                return true;
            }

            return false;
        }
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
                DataBase.ChangeAppSettings("DataAttiva", dataAttiva.ToString("yyyyMMdd"));
                return true;
            }
            return false;
        }

        public static void InitLog()
        {
            DataTable dtLog = DataBase.Select(DataBase.SP.APPLICAZIONE_LOG);
            dtLog.TableName = DataBase.Tab.LOG;
            if (DataBase.LocalDB.Tables.Contains(DataBase.Tab.LOG))
                DataBase.LocalDB.Tables[DataBase.Tab.LOG].Merge(dtLog);
            else
                DataBase.LocalDB.Tables.Add(dtLog);

            DataView dv = DataBase.LocalDB.Tables[DataBase.Tab.LOG].DefaultView;
            dv.Sort = "Data DESC";
        }
        private static int InitUser()
        {
            try
            {
                DataBase.ResetTable(DataBase.Tab.UTENTE);

                DataTable dtUtente = DataBase.Select(DataBase.SP.UTENTE, new QryParams() { { "@CodUtenteWindows", Environment.UserName } });
                dtUtente.TableName = DataBase.Tab.UTENTE;

                if (dtUtente.Rows.Count == 0)
                {
                    DataRow r = dtUtente.NewRow();
                    r["IdUtente"] = 0;
                    r["Nome"] = "NON CONFIGURATO";
                    dtUtente.Rows.Add(r);
                }
                DataBase.LocalDB.Tables.Add(dtUtente);

                return int.Parse("" + dtUtente.Rows[0]["IdUtente"]);
            }
            catch (Exception e)
            {
                DataBase.DB.Insert(DataBase.SP.INSERT_LOG, new QryParams() { { "@IdTipologia", Core.DataBase.TipologiaLOG.LogErrore }, { "@Messaggio", "InitUser: " + e.Message } });

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                return -1;
            }
        }
        private static bool Init(string dbName, string appID, DateTime dataAttiva)
        {
            //CryptHelper.CryptSection("connectionStrings", "appSettings");

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
            if (ConfigurationManager.AppSettings["Mercati"] != null)
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
        public static void AggiornaLabelStatoDB()
        {
            bool isProtected = true;
            try
            {
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
            catch
            { }
        }
        public static void DumpDataSet()
        {
            StringWriter strWriter = new StringWriter();
            XmlWriter xmlWriter = XmlWriter.Create(strWriter);
            Utility.DataBase.LocalDB.WriteXml(xmlWriter, XmlWriteMode.WriteSchema);

            XElement root = XElement.Parse(strWriter.ToString());
            XNamespace ns = WB.Name;//Simboli.NameSpace;

            IEnumerable<XElement> log =
                from tables in root.Elements(ns + Utility.DataBase.Tab.LOG)
                select tables;

            log.Remove();

            string locDBXml = strWriter.ToString();
            Microsoft.Office.Core.CustomXMLPart part;

            try
            {
                _wb.CustomXMLParts[WB.Name].Delete();
            }
            catch
            {
            }
            part = _wb.CustomXMLParts.Add();

            part.LoadXML(root.ToString(SaveOptions.DisableFormatting));
            //part.LoadXML(locDBXml);
        }

        public static void _db_StatoDBChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            AggiornaLabelStatoDB();
        }

        public static UserConfigElement GetUsrConfigElement(string configKey)
        {
            var settings = (UserConfiguration)ConfigurationManager.GetSection("usrConfig");

            return (UserConfigElement)settings.Items[configKey];
        }
        public static string AppSettings(string key)
        {
            try
            {
                return ConfigurationManager.AppSettings[key];
            }
            catch
            {
                ConfigurationManager.RefreshSection("applicationSettings");
                return ConfigurationManager.AppSettings[key];
            }
        }

        public static int[] GetRGBFromString(string rgb)
        {
            string[] rgbComp = rgb.Split(';');

            return new int[] { int.Parse(rgbComp[0]), int.Parse(rgbComp[1]), int.Parse(rgbComp[2]) };
        }
        
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
        }

        #endregion
    }   
}
