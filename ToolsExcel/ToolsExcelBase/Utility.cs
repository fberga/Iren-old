﻿using Iren.ToolsExcel.Base;
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
                APPLICAZIONE_INFORMAZIONE_COMMENTO = "spApplicazioneInformazioneCommento",
                APPLICAZIONE_INIT = "spApplicazioneInit",
                APPLICAZIONE_LOG = "spApplicazioneLog",
                APPLICAZIONE_NOTE = "spApplicazioneNote",
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
                INSERT_APPLICAZIONE_INFORMAZIONE_D = "spInsertApplicazioneInformazioneD",
                INSERT_APPLICAZIONE_RIEPILOGO = "spInsertApplicazioneRiepilogo",
                INSERT_LOG = "spInsertLog",
                INSERT_PROGRAMMAZIONE_PARAMETRO = "spInsertProgrammazione_Parametro",
                TIPOLOGIA_CHECK = "spTipologiaCheck",
                TIPOLOGIA_RAMPA = "spTipologiaRampa",
                UTENTE = "spUtente";
        }
        public struct Tab
        {
            public const string ADDRESS_FROM = "AddressFrom",
                ADDRESS_TO = "AddressTo",
                ANNOTA = "AnnotaModifica",
                APPLICAZIONE = "Applicazione",
                AZIONE = "Azione",
                AZIONE_CATEGORIA = "AzioneCategoria",
                CALCOLO = "Calcolo",
                CALCOLO_INFORMAZIONE = "CalcoloInformazione",
                CATEGORIA = "Categoria",
                CATEGORIA_ENTITA = "CategoriaEntita",
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
                NOMI_DEFINITI_NEW = "DefinedNamesNew",
                SALVADB = "SaveDB",
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
        public static void RefreshDate(DateTime dataAttiva) 
        {
            _db.ChangeDate(dataAttiva.ToString("yyyyMMdd"));
        }
        public static void RefreshAppSettings(string key, string value) 
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings[key].Value = value;
            config.Save(ConfigurationSaveMode.Minimal);
            ConfigurationManager.RefreshSection("appSettings");
        }
        public static void SwitchEnvironment(string ambiente) 
        {
            RefreshAppSettings("DB", ambiente);
            ConfigurationManager.RefreshSection("appSettings");

            _db = new Core.DataBase(ambiente);
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

                var path = Utilities.GetUsrConfigElement("pathExportModifiche");

                string cartellaRemota = ExportPath.PreparePath(path.Value);
                string cartellaEmergenza = ExportPath.PreparePath(path.Emergenza);
                string cartellaArchivio = ExportPath.PreparePath(path.Archivio);

                string fileName = "";
                if (onLine && Directory.Exists(cartellaRemota))
                {
                    string[] fileEmergenza = Directory.GetFiles(cartellaEmergenza);

                    if (fileEmergenza.Length > 0)
                    {
                        Array.Sort<string>(fileEmergenza);
                        foreach (string file in fileEmergenza)
                        {
                            File.Move(file, Path.Combine(cartellaRemota, file.Split('\\').Last()));
                            //TODO esegui stored procedure sul file
                            if (true)
                                File.Move(Path.Combine(cartellaRemota, file.Split('\\').Last()), Path.Combine(cartellaArchivio, file.Split('\\').Last()));
                        }
                    }

                    fileName = Path.Combine(cartellaRemota, Simboli.nomeApplicazione.Replace(" ", "").ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml");
                    dt.WriteXml(fileName);
                    //TODO esegui stored procedure
                    if (true)
                        File.Move(fileName, Path.Combine(cartellaArchivio, fileName.Split('\\').Last()));
                }
                else
                {
                    fileName = Path.Combine(cartellaEmergenza, Simboli.nomeApplicazione.Replace(" ", "").ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml");
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
                DB.CloseConnection();
            }
        }
        public static object GetMessaggioCheck(object id) 
        {
            DataView tipologiaCheck = _localDB.Tables[Tab.TIPOLOGIA_CHECK].DefaultView;
            tipologiaCheck.RowFilter = "IdTipologiaCheck = " + id;

            if (tipologiaCheck.Count > 0)
                return tipologiaCheck[0]["Messaggio"];

            return null;
        }
        public static void InsertLog(Core.DataBase.TipologiaLOG logType, string message) 
        {
            #if (!DEBUG)
            if (_db.OpenConnection())
            {
                _db.Insert(SP.INSERT_LOG, new QryParams() { { "@IdTipologia", logType }, { "@Messaggio", message } });
                DataTable dt = _db.Select(SP.APPLICAZIONE_LOG);
                if (_localDB.Tables[Tab.LOG] != null && dt.Columns != _localDB.Tables[Tab.LOG].Columns)
                    ResetTable(Tab.LOG);

                dt.TableName = Tab.LOG;
                _localDB.Merge(dt);
            }
            #endif
        }
        public static void InsertApplicazioneRiepilogo(object siglaEntita, object siglaAzione, DateTime? dataRif = null, bool presente = true) 
        {
            dataRif = dataRif ?? DataAttiva;
            try
            {
                _db.OpenConnection();
                QryParams parameters = new QryParams() {
                    {"@SiglaEntita", siglaEntita},
                    {"@SiglaAzione", siglaAzione},
                    {"@Data", dataRif.Value.ToString("yyyyMMdd")},
                    {"@Presente", presente ? "1" : "0"}
                };
                _db.Insert(DataBase.SP.INSERT_APPLICAZIONE_RIEPILOGO, parameters);
            }
            catch (Exception e)
            {
                InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "InsertApplicazioneRiepilogo [" + dataRif ?? DataAttiva + ", " + siglaEntita + ", " + siglaAzione + "]: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        public static void ConvertiParametriInformazioni() 
        {
            _db.Select(SP.APPLICAZIONE_INIT);
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

        #endregion
    }

    public class ExportPath 
    {
        #region Metodi

        public static string PreparePath(string path) 
        {
            Regex options = new Regex(@"\[\w+\]");
            path = options.Replace(path, match =>
            {
                string opt = match.Value.Replace("[", "").Replace("]", "");
                string o = "";
                switch (opt.ToLowerInvariant())
                {
                    case "appname":
                        o = Simboli.nomeApplicazione.Replace(" ", "");
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

        public static string GetSuffissoDATA1
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
        #endregion
    }

    public class Struttura : DataBase
    {
        #region Metodi

        public static void AggiornaParametriApplicazione(object appID)
        {
            DataTable dt = CaricaApplicazione(appID);
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
        private static DataTable CaricaApplicazione(object idApplicazione)
        {
            string name = DataBase.Tab.APPLICAZIONE;
            DataBase.ResetTable(name);
            QryParams parameters = new QryParams() 
            {
                {"@IdApplicazione", idApplicazione},

            };
            DataTable dt = DataBase.DB.Select(DataBase.SP.APPLICAZIONE, parameters);
            dt.TableName = name;
            return dt;
        }

        public static void AggiornaStrutturaDati()
        {
            //CreaTabellaNomi();
            CreaTabellaNomiNew();
            CreaTabellaDate();
            CreaTabellaAddressFrom();
            CreaTabellaAddressTo();
            CreaTabellaModifica();
            CreaTabellaEditabili();
            CreaTabellaSalvaDB();
            CreaTabellaAnnotaModifica();
            CaricaAzioni();
            CaricaCategorie();
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

        //private static bool CreaTabellaNomi()
        //{
        //    try
        //    {
        //        string name = Tab.NOMI_DEFINITI;
        //        ResetTable(name);
        //        DataTable dt = DefinedNames.GetDefaultTable(name);
        //        _localDB.Tables.Add(dt);
        //        return true;
        //    }
        //    catch (Exception)
        //    {
        //        return false;
        //    }
        //}

        private static bool CreaTabellaNomiNew()
        {
            try
            {
                string name = Tab.NOMI_DEFINITI_NEW;
                ResetTable(name);
                DataTable dt = NewDefinedNames.GetDefaultNameTable(name);
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
                DataTable dt = NewDefinedNames.GetDefaultDateTable(name);
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
                DataTable dt = NewDefinedNames.GetDefaultAddressFromTable(name);
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
                DataTable dt = NewDefinedNames.GetDefaultAddressToTable(name);
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
                DataTable dt = NewDefinedNames.GetDefaultEditableTable(name);
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
                DataTable dt = NewDefinedNames.GetDefaultSaveTable(name);
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
                DataTable dt = NewDefinedNames.GetDefaultToNote(name);
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
                DataTable dt = _db.Select(SP.AZIONE, parameters);
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
                DataTable dt = _db.Select(SP.AZIONE, parameters);
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
                DataTable dt = _db.Select(SP.CATEGORIA, parameters);
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
                DataTable dt = _db.Select(SP.AZIONE_CATEGORIA, parameters);
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
                DataTable dt = _db.Select(SP.CATEGORIA_ENTITA, parameters);
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
                DataTable dt = _db.Select(SP.ENTITA_AZIONE, parameters);
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
                DataTable dt = _db.Select(SP.ENTITA_AZIONE_CALCOLO, parameters);
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
                DataTable dt = _db.Select(SP.ENTITA_INFORMAZIONE, parameters);
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
                DataTable dt = _db.Select(SP.ENTITA_AZIONE_INFORMAZIONE, parameters);
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
                DataTable dt = _db.Select(SP.CALCOLO, parameters);
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
                DataTable dt = _db.Select(SP.CALCOLO_INFORMAZIONE, parameters);
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
                DataTable dt = _db.Select(SP.ENTITA_CALCOLO, parameters);
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
                DataTable dt = _db.Select(SP.ENTITA_GRAFICO, parameters);
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
                DataTable dt = _db.Select(SP.ENTITA_GRAFICO_INFORMAZIONE, parameters);
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
                DataTable dt = _db.Select(SP.ENTITA_COMMITMENT, parameters);
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
                DataTable dt = _db.Select(SP.ENTITA_RAMPA, parameters);
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
                DataTable dt = _db.Select(SP.ENTITA_ASSETTO, parameters);
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
                DataTable dt = _db.Select(SP.ENTITA_PROPRIETA, parameters);
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
                DataTable dt = _db.Select(SP.ENTITA_INFORMAZIONE_FORMATTAZIONE, parameters);
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
                DataTable dt = _db.Select(SP.TIPOLOGIA_CHECK, parameters);
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
                DataTable dt = _db.Select(SP.TIPOLOGIA_RAMPA, parameters);
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
                DataTable dt = _db.Select(SP.ENTITA_PARAMETRO_D);
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
                DataTable dt = _db.Select(SP.ENTITA_PARAMETRO_H);
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

        #endregion
    }

    public class Workbook 
    {
        #region Variabili

        protected static Microsoft.Office.Tools.Excel.Workbook _wb;
        private static bool _daElaborazione = false;

        #endregion

        #region Proprietà

        public static Microsoft.Office.Tools.Excel.Workbook WB { get { return _wb; } }
        public static bool DaElaborazione
        {
            get { return _daElaborazione; }
        }

        #endregion

        #region Metodi

        public static void AggiornaFormule(Excel.Worksheet ws)
        {
            ws.Application.CalculateFull();
        }
        public static bool CaricaAzioneInformazione(object siglaEntita, object siglaAzione, object azionePadre, DateTime giorno, object parametro = null)
        {
            NewDefinedNames newNomiDefiniti = new NewDefinedNames(NewDefinedNames.GetSheetName(siglaEntita));

            try
            {
                DataView azioni = DataBase.LocalDB.Tables[DataBase.Tab.AZIONE].DefaultView;
                azioni.RowFilter = "SiglaAzione = '" + siglaAzione + "'";

                //if (dataRif == null)
                //    dataRif = DataBase.DataAttiva;

                bool procedi = true;
                if (azioni[0]["Visibile"].Equals("1"))
                {
                    procedi = false;
                    DataView entitaAzioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE].DefaultView;
                    entitaAzioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";
                    if (entitaAzioni.Count > 0)
                    {
                        procedi = true;
                    }
                }

                if (procedi)
                {
                    AzzeraInformazione2(siglaEntita, siglaAzione, newNomiDefiniti, giorno);

                    if (DataBase.OpenConnection())
                    {
                        if (azionePadre.Equals("GENERA"))
                        {
                            ElaborazioneInformazione2(siglaEntita, siglaAzione, newNomiDefiniti, giorno, (siglaAzione.Equals("G_MP_MGP") ? 5 : 7));
                            if (azioni[0]["Visibile"].Equals("1"))
                                DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, giorno);
                        }
                        else
                        {
                            DataView azioneInformazione = DataBase.DB.Select(DataBase.SP.CARICA_AZIONE_INFORMAZIONE, "@SiglaEntita=" + siglaEntita + ";@SiglaAzione=" + siglaAzione + ";@Parametro=" + parametro + ";@Data=" + giorno.ToString("yyyyMMdd")).DefaultView;
                            if (azioneInformazione.Count == 0)
                            {
                    //            if (azioni[0]["Visibile"].Equals("1"))
                    //                DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, dataRif, false);
                            }
                            else
                            {
                                ScriviInformazione2(siglaEntita, azioneInformazione, newNomiDefiniti);

                                //if (azioni[0]["Visibile"].Equals("1"))
                                //    DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, giorno);
                            }
                        }
                    }
                    else
                    {
                        if (azionePadre.Equals("GENERA"))
                            ElaborazioneInformazione2(siglaEntita, siglaAzione, newNomiDefiniti, giorno, (siglaAzione.Equals("G_MP_MGP") ? 5 : 7));
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "CaricaAzioneInformazione [" + siglaEntita + ", " + siglaAzione + "]: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        public static void AzzeraInformazione2(object siglaEntita, object siglaAzione, NewDefinedNames newNomiDefiniti, DateTime giorno)
        {
            Excel.Worksheet ws = _wb.Sheets[newNomiDefiniti.Sheet];

            string suffissoData = Date.GetSuffissoData(giorno);

            DataView azioneInformazione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
            azioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";

            foreach (DataRowView info in azioneInformazione)
            {
                if (info["FormulaInCella"].Equals("0"))
                {
                    siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                    Range rng;
                    if(info["Selezione"].Equals(0))
                        rng = newNomiDefiniti.Get(siglaEntita, info["SiglaInformazione"], suffissoData).Extend(colOffset: Date.GetOreGiorno(giorno));
                    else
                        rng = newNomiDefiniti.Get(siglaEntita, "SEL", info["SiglaInformazione"], suffissoData).Extend(colOffset: Date.GetOreGiorno(giorno));

                    Excel.Range xlRng = ws.Range[rng.ToString()];
                    xlRng.Value = null;
                    Style.RangeStyle(xlRng, backColor: info["BackColor"], foreColor: info["ForeColor"]);
                    xlRng.ClearComments();
                }
            }
        }
        public static void ScriviInformazione2(object siglaEntita, DataView azioneInformazione, NewDefinedNames newNomiDefiniti)
        {
            Excel.Worksheet ws = _wb.Sheets[newNomiDefiniti.Sheet];

            foreach (DataRowView azione in azioneInformazione)
            {
                string suffissoData;
                string suffissoOra;
                if (azione["SiglaEntita"].Equals("UP_BUS") && azione["SiglaInformazione"].Equals("VOL_INVASO")) 
                {
                    suffissoData = Date.GetSuffissoData(DataBase.DataAttiva.AddDays(-1));
                    suffissoOra = Date.GetSuffissoOra(24);
                }
                else
                {
                    suffissoData = Date.GetSuffissoData(DataBase.DataAttiva, azione["Data"]);
                    suffissoOra = Date.GetSuffissoOra(azione["Data"]);
                }

                Range rng = newNomiDefiniti.Get(siglaEntita, azione["SiglaInformazione"], suffissoData, suffissoOra);

                Excel.Range xlRng = ws.Range[rng.ToString()];
                xlRng.Value = azione["Valore"];
                if (azione["BackColor"] != DBNull.Value)
                    xlRng.Interior.ColorIndex = azione["BackColor"];
                if (azione["BackColor"] != DBNull.Value)
                    xlRng.Font.ColorIndex = azione["ForeColor"];

                xlRng.ClearComments();

                if (azione["Commento"] != DBNull.Value)
                    xlRng.AddComment(azione["Commento"]).Visible = false;
            }
        }
        private static void ElaborazioneInformazione2(object siglaEntita, object siglaAzione, NewDefinedNames newNomiDefiniti, DateTime giorno, int tipologiaCalcolo, int oraInizio = -1, int oraFine = -1)
        {
            _daElaborazione = true;
            Excel.Worksheet ws = _wb.Sheets[newNomiDefiniti.Sheet];

            Dictionary<string, int> entitaRiferimento = new Dictionary<string, int>();
            List<int> oreDaCalcolare = new List<int>();

            string suffissoData = Date.GetSuffissoData(giorno);
            
            oraInizio = oraInizio < 0 ? 1 : oraInizio;
            oraFine = oraFine < 0 ? Date.GetOreGiorno(giorno) : oraFine;

            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "Gerarchia = '" + siglaEntita + "'";
            foreach (DataRowView entita in categoriaEntita)
                entitaRiferimento.Add(entita["SiglaEntita"].ToString(), (int)entita["Riferimento"]);

            if (entitaRiferimento.Count == 0)
                entitaRiferimento.Add(siglaEntita.ToString(), 1);

            DataView calcoloInformazione = DataBase.LocalDB.Tables[DataBase.Tab.CALCOLO_INFORMAZIONE].DefaultView;

            DataView entitaAzioneCalcolo = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE_CALCOLO].DefaultView;
            entitaAzioneCalcolo.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";
            foreach (DataRowView azioneCalcolo in entitaAzioneCalcolo)
            {
                calcoloInformazione.RowFilter = "SiglaCalcolo = '" + azioneCalcolo["SiglaCalcolo"] + "'";
                calcoloInformazione.Sort = "Step";

                //azzero tutte le informazioni che vengono utilizzate nel calcolo tranne i CHECK
                foreach (DataRowView info in calcoloInformazione)
                {
                    if (!info["SiglaInformazione"].Equals("CHECKINFO"))
                    {
                        Range rng = newNomiDefiniti.Get(info["SiglaEntitaRif"] is DBNull ? siglaEntita : info["SiglaEntitaRif"], info["SiglaInformazione"], suffissoData, Date.GetSuffissoOra(oraInizio)).Extend(colOffset: oraFine - oraInizio + 1);
                        ws.Range[rng.ToString()].Value = null;
                    }
                }

                for (int ora = oraInizio; ora <= oraFine; ora++)
                {
                    int i = 0;
                    while (i < calcoloInformazione.Count)
                    {
                        DataRowView calcolo = calcoloInformazione[i];
                        if (calcolo["OraInizio"] != DBNull.Value)
                            if (ora < int.Parse(calcolo["OraInizio"].ToString()) || ora > int.Parse(calcolo["OraFine"].ToString()))
                                continue;

                        if (calcolo["OraFine"] != DBNull.Value)
                            if (ora != Date.GetOreGiorno(giorno))
                                if (calcolo["FineCalcolo"].Equals("1"))
                                    continue;
                                else
                                    break;

                        int step = 0;
                        object risultato = GetRisultatoCalcolo2(siglaEntita, newNomiDefiniti, giorno, ora, calcolo, entitaRiferimento, out step);

                        if (step == 0)
                        {
                            Range rng = newNomiDefiniti.Get(siglaEntita, calcolo["SiglaInformazione"], suffissoData, Date.GetSuffissoOra(ora));
                            Excel.Range xlRng = ws.Range[rng.ToString()];

                            xlRng.Formula = calcolo["SiglaInformazione"].Equals("CHECKINFO") ? DataBase.GetMessaggioCheck(risultato) : risultato;

                            if (calcolo["BackColor"] != DBNull.Value)
                                xlRng.Interior.ColorIndex = calcolo["BackColor"];
                            if (calcolo["ForeColor"] != DBNull.Value)
                                xlRng.Font.ColorIndex = calcolo["ForeColor"];

                            xlRng.ClearComments();

                            if (calcolo["Commento"] != DBNull.Value)
                                xlRng.AddComment(calcolo["Commento"]).Visible = false;

                            Handler.StoreEdit(ws, xlRng);
                        }

                        if (calcolo["FineCalcolo"].Equals("1") || step == -1)
                            break;

                        if (calcolo["GoStep"] != DBNull.Value)
                            step = (int)calcolo["GoStep"];

                        if (step != 0)
                            i = calcoloInformazione.Find(step);
                        else
                            i++;
                    }
                }
            }
            _daElaborazione = false;
        }
        private static object GetRisultatoCalcolo2(object siglaEntita, NewDefinedNames newNomiDefiniti, DateTime giorno, int ora, DataRowView calcolo, Dictionary<string, int> entitaRiferimento, out int step)
        {
            Excel.Worksheet ws = _wb.Sheets[newNomiDefiniti.Sheet];

            string suffissoData = Date.GetSuffissoData(giorno);

            int ora1 = calcolo["OraInformazione1"] is DBNull ? ora : ora + (int)calcolo["OraInformazione1"];
            int ora2 = calcolo["OraInformazione2"] is DBNull ? ora : ora + (int)calcolo["OraInformazione2"];

            object siglaEntitaRif1 = calcolo["Riferimento1"] is DBNull ? (calcolo["SiglaEntita1"] is DBNull ? siglaEntita : calcolo["SiglaEntita1"]) : entitaRiferimento.FirstOrDefault(kv => kv.Value == (int)calcolo["Riferimento1"]);
            object siglaEntitaRif2 = calcolo["Riferimento2"] is DBNull ? (calcolo["SiglaEntita2"] is DBNull ? siglaEntita : calcolo["SiglaEntita2"]) : entitaRiferimento.FirstOrDefault(kv => kv.Value == (int)calcolo["Riferimento2"]);

            object valore1 = 0d;
            object valore2 = 0d;

            if (calcolo["SiglaInformazione1"] != DBNull.Value)
            {
                Range cella1 = newNomiDefiniti.Get(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1));

                switch (calcolo["SiglaInformazione1"].ToString())
                {
                    case "UNIT_COMM":
                        DataView entitaCommitment = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_COMMITMENT].DefaultView;
                        entitaCommitment.RowFilter = "SiglaCommitment = '" + ws.Range[cella1.ToString()].Value + "'";
                        valore1 = entitaCommitment.Count > 0 ? entitaCommitment[0]["IdEntitaCommitment"] : null;

                        break;
                    case "DISPONIBILITA":
                        if (ws.Range[cella1.ToString()].Value == "OFF")
                            valore1 = 0d;
                        else
                            valore1 = 1d;

                        break;
                    case "CHECKINFO":
                        if (ws.Range[cella1.ToString()].Value == "OK")
                            valore1 = 1d;
                        else
                            valore1 = 2d;
                        break;
                    default:
                        //if (cella != null)
                            valore1 = ws.Range[cella1.ToString()].Value ?? 0d;
                        break;
                }
            }
            else if (calcolo["IdProprieta"] != DBNull.Value)
            {
                DataView entitaProprieta = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' IdProprieta = " + calcolo["IdProprieta"];

                if (entitaProprieta.Count > 0)
                    valore1 = entitaProprieta[0]["Valore"];
            }
            else if (calcolo["IdParametroD"] != DBNull.Value)
            {
                DataView entitaParametro = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PARAMETRO_D].DefaultView;
                entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' IdParametro = " + calcolo["IdParametroD"];

                if (entitaParametro.Count > 0)
                    valore1 = entitaParametro[0]["Valore"].ToString().Replace('.', ',');
            }
            else if (calcolo["IdParametroH"] != DBNull.Value)
            {
                DataView entitaParametro = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PARAMETRO_H].DefaultView;
                entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' IdParametro = " + calcolo["IdParametroH"];

                if (entitaParametro.Count > 0)
                    valore1 = entitaParametro[0]["Valore"].ToString().Replace('.', ',');
            }
            else if (calcolo["Valore"] != DBNull.Value)
            {
                valore1 = Convert.ToDouble(calcolo["Valore"]);
            }

            if (calcolo["SiglaInformazione2"] != DBNull.Value)
            {
                Range cella2 = newNomiDefiniti.Get(siglaEntitaRif1, calcolo["SiglaInformazione2"], suffissoData, Date.GetSuffissoOra(ora2));

                switch (calcolo["SiglaInformazione2"].ToString())
                {
                    case "UNIT_COMM":
                        DataView entitaCommitment = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_COMMITMENT].DefaultView;
                        entitaCommitment.RowFilter = "SiglaCommitment = '" + ws.Range[cella2.ToString()].Value + "'";
                        valore2 = entitaCommitment.Count > 0 ? entitaCommitment[0] : null;

                        break;
                    case "DISPONIBILITA":
                        if (ws.Range[cella2.ToString()].Value == "OFF")
                            valore2 = 0d;
                        else
                            valore2 = 1d;

                        break;
                    case "CHECKINFO":
                        if (ws.Range[cella2.ToString()].Value == "OK")
                            valore2 = 1d;
                        else
                            valore2 = 2d;
                        break;
                    default:
                        //if (cella != null)
                        valore2 = ws.Range[cella2.ToString()].Value ?? 0d;
                        //else
                        //    valore2 = 0d;
                        break;
                }
            }

            double retVal = 0d;

            valore1 = valore1 ?? 0d;
            valore2 = valore2 ?? 0d;

            if (calcolo["Funzione"] is DBNull && calcolo["Operazione"] is DBNull && calcolo["Condizione"] is DBNull)
            {
                step = 0;
                if (Convert.ToDouble(valore1) == 0d)
                    return valore2;

                return valore1;
            }
            else if (calcolo["Funzione"] != DBNull.Value)
            {
                string func = calcolo["Funzione"].ToString().ToLowerInvariant();
                if (calcolo["SiglaInformazione2"] is DBNull)
                {
                    if (func.Contains("abs"))
                    {
                        retVal = Math.Abs(Convert.ToDouble(valore1));
                    }
                    else if (func.Contains("floor"))
                    {
                        retVal = Math.Floor(Convert.ToDouble(valore1));
                    }
                    else if (func.Contains("round"))
                    {
                        int decimals = int.Parse(func.Replace("round", ""));
                        retVal = Math.Round(Convert.ToDouble(valore1), decimals);
                    }
                    else if (func.Contains("power"))
                    {
                        int exp = int.Parse(Regex.Match(func, @"\d*").Value);
                        retVal = Math.Pow(Convert.ToDouble(valore1), exp);
                    }
                    else if (func.Contains("sum"))
                    {
                        foreach (var kvp in entitaRiferimento)
                            retVal += ws.Range[newNomiDefiniti.Get(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1)).ToString()].Value ?? 0d;
                    }
                    else if (func.Contains("avg"))
                    {
                        foreach (var kvp in entitaRiferimento)
                            retVal += ws.Range[newNomiDefiniti.Get(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1)).ToString()].Value ?? 0d;
                        retVal /= entitaRiferimento.Count;
                    }
                    else if (func.Contains("max_h"))
                    {                        
                        Range rng = newNomiDefiniti.Get(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData).Extend(colOffset: Date.GetOreGiorno(giorno));
                        object[,] tmpVal = ws.Range[rng.ToString()].Value;
                        double[] values = tmpVal.Cast<double>().ToArray();
                        retVal = values.Max();
                    }
                    else if (func.Contains("min_h"))
                    {
                        Range rng = newNomiDefiniti.Get(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData).Extend(colOffset: Date.GetOreGiorno(giorno));
                        object[,] tmpVal = ws.Range[rng.ToString()].Value;
                        double[] values = tmpVal.Cast<double>().ToArray();
                        retVal = values.Min();
                    }
                    else if (func.Contains("max"))
                    {
                        retVal = double.MinValue;
                        foreach (var kvp in entitaRiferimento)
                            retVal = Math.Max(ws.Range[newNomiDefiniti.Get(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1)).ToString()].Value ?? 0, retVal);
                    }
                    else if (func.Contains("min"))
                    {
                        retVal = double.MaxValue;
                        foreach (var kvp in entitaRiferimento)
                            retVal = Math.Min(ws.Range[newNomiDefiniti.Get(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1)).ToString()].Value ?? 0, retVal);
                    }
                }
                //caso in cui ci sia anche SiglaInformazione2
                else
                {
                    if (func.Contains("max"))
                    {
                        retVal = Math.Max(Convert.ToDouble(valore1), Convert.ToDouble(valore2));
                    }
                    else if (func.Contains("min"))
                    {
                        retVal = Math.Min(Convert.ToDouble(valore1), Convert.ToDouble(valore2));
                    }
                }
            }
            else if (calcolo["Operazione"] != DBNull.Value)
            {
                switch (calcolo["Operazione"].ToString())
                {
                    case "+":
                        retVal = Convert.ToDouble(valore1) + Convert.ToDouble(valore2);
                        break;
                    case "-":
                        retVal = Convert.ToDouble(valore1) - Convert.ToDouble(valore2);
                        break;
                    case "*":
                        retVal = Convert.ToDouble(valore1) * Convert.ToDouble(valore2);
                        break;
                    case "/":
                        retVal = Convert.ToDouble(valore1) / Convert.ToDouble(valore2);
                        break;
                }
            }
            else if (calcolo["Condizione"] != DBNull.Value)
            {
                bool res = false;
                switch (calcolo["Condizione"].ToString())
                {
                    case ">":
                        res = Convert.ToDouble(valore1) > Convert.ToDouble(valore2);
                        break;
                    case "<":
                        res = Convert.ToDouble(valore1) < Convert.ToDouble(valore2);
                        break;
                    case ">=":
                        res = Convert.ToDouble(valore1) >= Convert.ToDouble(valore2);
                        break;
                    case "<=":
                        res = Convert.ToDouble(valore1) <= Convert.ToDouble(valore2);
                        break;
                    case "=":
                        res = Convert.ToDouble(valore1) == Convert.ToDouble(valore2);
                        break;
                    case "<>":
                        res = Convert.ToDouble(valore1) != Convert.ToDouble(valore2);
                        break;
                }
                if (res)
                    step = (int)calcolo["StepCondizioneVera"];
                else
                    step = (int)calcolo["StepCondizioneFalsa"];

                return res;
            }

            step = 0;
            return retVal;
        }

        //public static void AzzeraInformazione(object siglaEntita, object siglaAzione, DateTime? dataRif = null, object valore = null)
        //{
        //    string foglio = DefinedNames.GetSheetName(siglaEntita);

        //    DefinedNames nomiDefiniti = new DefinedNames(foglio);
        //    Excel.Worksheet ws = _wb.Sheets.OfType<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name == foglio);

        //    if (dataRif == null)
        //        dataRif = DataBase.DataAttiva;

        //    string suffissoData = Date.GetSuffissoData(DataBase.DataAttiva, dataRif.Value);

        //    DataView entitaAzioniInformazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE_INFORMAZIONE].DefaultView;

        //    //TODO controllare perché Domenico passa un entitaRif a true/false in questo filtro
        //    entitaAzioniInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";

        //    foreach (DataRowView entitaAzioneInformazione in entitaAzioniInformazioni)
        //    {
        //        if (entitaAzioneInformazione["FormulaInCella"].Equals("0"))
        //        {
        //            object entita = entitaAzioneInformazione["SiglaEntitaRif"] is DBNull ? entitaAzioneInformazione["SiglaEntita"] : entitaAzioneInformazione["SiglaEntitaRif"];
        //            Tuple<int, int>[] riga = new Tuple<int, int>[0];

        //            if (entitaAzioneInformazione["Selezione"].Equals(0))
        //                riga = nomiDefiniti[DefinedNames.GetName(entita, entitaAzioneInformazione["SiglaInformazione"], suffissoData)];
        //            else
        //                riga = nomiDefiniti[DefinedNames.GetName(entita, "SEL", entitaAzioneInformazione["Selezione"], suffissoData)];

        //            Excel.Range rng = ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]];
        //            rng.Value = valore;
        //            rng.Interior.ColorIndex = entitaAzioneInformazione["BackColor"];
        //            rng.Font.ColorIndex = entitaAzioneInformazione["ForeColor"];
        //            rng.ClearComments();
        //        }
        //    }
        //}
        //public static void ScriviInformazione(object siglaEntita, DataView azioneInformazione)
        //{
        //    string foglio = DefinedNames.GetSheetName(siglaEntita);

        //    DefinedNames nomiDefiniti = new DefinedNames(foglio);
        //    Excel.Worksheet ws = _wb.Sheets[foglio];

        //    foreach (DataRowView azione in azioneInformazione)
        //    {
        //        string suffissoData;
        //        if (azione["SiglaEntita"].Equals("UP_BUS") && azione["SiglaInformazione"].Equals("VOL_INVASO"))
        //            suffissoData = DefinedNames.GetName("DATA0", "H24");
        //        else
        //            suffissoData = DefinedNames.GetName(Date.GetSuffissoData(DataBase.DataAttiva, azione["Data"]), Date.GetSuffissoOra(azione["Data"]));

        //        Tuple<int, int>[] celle = nomiDefiniti[DefinedNames.GetName(azione["SiglaEntita"], azione["SiglaInformazione"], suffissoData)];
        //        if (celle != null)
        //        {
        //            Excel.Range rng = ws.Cells[celle[0].Item1, celle[0].Item2];
        //            rng.Value = azione["Valore"];
        //            if (azione["BackColor"] != DBNull.Value)
        //                rng.Interior.ColorIndex = azione["BackColor"];
        //            if (azione["BackColor"] != DBNull.Value)
        //                rng.Font.ColorIndex = azione["ForeColor"];

        //            rng.ClearComments();

        //            if (azione["Commento"] != DBNull.Value)
        //                rng.AddComment(azione["Commento"]).Visible = false;
        //        }
        //    }
        //}
        //private static void ElaborazioneInformazione(object siglaEntita, DateTime data, int tipologiaCalcolo, int oraInizio = 0, int oraFine = 0)
        //{
        //    Dictionary<object, object> entitaRiferimento = new Dictionary<object, object>();
        //    List<int> oreDaCalcolare = new List<int>();


        //    string suffissoData = Date.GetSuffissoData(DataBase.DataAttiva, data);
        //    string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
        //    Excel.Worksheet ws = _wb.Sheets.OfType<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name == nomeFoglio);
        //    DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);

        //    if (oraInizio == 0)
        //    {
        //        oraInizio++;
        //        oraFine = Date.GetOreGiorno(data);
        //    }
        //    DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
        //    DataView entitaInformazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
        //    categoriaEntita.RowFilter = "Gerarchia = '" + siglaEntita + "'";
        //    foreach (DataRowView entita in categoriaEntita)
        //        entitaRiferimento.Add(entita["SiglaEntita"].ToString(), entita["Riferimento"]);

        //    if (entitaRiferimento.Count == 0)
        //        entitaRiferimento.Add(siglaEntita, 1);


        //    if (tipologiaCalcolo == 1 || tipologiaCalcolo == 5 && DefinedNames.IsDefined(nomeFoglio, DefinedNames.GetName(siglaEntita, "UNIT_COMM")))
        //    {
        //        DataView entitaCommitment = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_COMMITMENT].DefaultView;

        //        Tuple<int, int> primaCella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "UNIT_COMM", suffissoData, "H" + oraInizio)][0];
        //        Tuple<int, int> ultimaCella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "UNIT_COMM", suffissoData, "H" + oraFine)][0];
        //        object[,] tmpVal = ws.Range[ws.Cells[primaCella.Item1, primaCella.Item2], ws.Cells[ultimaCella.Item1, ultimaCella.Item2]].Value;
        //        object[] values = tmpVal.Cast<object>().ToArray();
        //        for (int i = oraInizio; i < oraFine; i++)
        //        {
        //            entitaCommitment.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaCommitment = '" + values[i - oraInizio] + "' AND AbilitaOffertaMGP = '1'";
        //            if (entitaCommitment.Count > 0)
        //                oreDaCalcolare.Add(i);
        //        }
        //    }
        //    else
        //    {
        //        for (int i = oraInizio; i < oraFine; i++)
        //            oreDaCalcolare.Add(i);
        //    }

        //    if (oreDaCalcolare.Count > 0)
        //    {
        //        if (tipologiaCalcolo == 3)
        //        {
        //            foreach (int ora in oreDaCalcolare)
        //            {
        //                Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "CHECKINFO", suffissoData, "H" + ora)][0];
        //                ws.Range[cella.Item1, cella.Item2].Value = null;
        //            }
        //        }

        //        DataView entitaCalcoli = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_CALCOLO].DefaultView;
        //        entitaCalcoli.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdTipologiaCalcolo = " + tipologiaCalcolo;

        //        DataView calcoloInformazioni = DataBase.LocalDB.Tables[DataBase.Tab.CALCOLO_INFORMAZIONE].DefaultView;

        //        foreach (DataRowView entitaCalcolo in entitaCalcoli)
        //        {
        //            calcoloInformazioni.RowFilter = "SiglaCalcolo = '" + entitaCalcolo["SiglaCalcolo"] + "'";
        //            DataView tmp = calcoloInformazioni.ToTable(true, "SiglaEntitaRif", "SiglaInformazione").DefaultView;

        //            foreach (int ora in oreDaCalcolare)
        //            {
        //                //foreach (DataRowView calcoloInfo in tmp)
        //                //{
        //                //    if (!calcoloInfo["SiglaInformazione"].Equals("CHECKINFO"))
        //                //    {
        //                //        object siglaEntita2 = calcoloInfo["SiglaEntitaRif"] is DBNull ? siglaEntita : calcoloInfo["SiglaEntitaRif"];
        //                //        if (DefinedNames.IsDefined(nomeFoglio, GetName(siglaEntita2, calcoloInfo["SiglaInformazione"], suffissoData, "H" + ora)))
        //                //        {
        //                //            Tuple<int, int> cella = nomiDefiniti[GetName(siglaEntita2, calcoloInfo["SiglaInformazione"], suffissoData, "H" + ora)][0];
        //                //            ws.Cells[cella.Item1, cella.Item2].Value = null;
        //                //        }
        //                //    }
        //                //}

        //                List<DataRowView> calcoloRows = calcoloInformazioni.Cast<DataRowView>().ToList();
        //                int i = 0;

        //                while (i < calcoloRows.Count)
        //                {
        //                    DataRowView calcolo = calcoloRows[i];

        //                    if (calcolo["OraInizio"] != DBNull.Value)
        //                        if (ora < oraInizio && ora > oraFine)
        //                        {
        //                            i++;
        //                            continue;
        //                        }


        //                    if (calcolo["OraFine"].Equals("0"))
        //                    {
        //                        if (ora != Date.GetOreGiorno(data))
        //                            break;
        //                        i++;
        //                        continue;
        //                    }

        //                    int step;
        //                    Stopwatch watch = Stopwatch.StartNew();
        //                    object risultato = GetRisultatoCalcolo(siglaEntita, data, ora, calcolo, entitaRiferimento, out step);
        //                    watch.Stop();
        //                    watch = Stopwatch.StartNew();
        //                    if (step == 0)
        //                    {
        //                        if (nomiDefiniti.IsDefined(DefinedNames.GetName(siglaEntita, calcolo["SiglaInformazione"], suffissoData, "H" + ora)))
        //                        {
        //                            Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, calcolo["SiglaInformazione"], suffissoData, "H" + ora)][0];

        //                            Excel.Range rng = ws.Cells[cella.Item1, cella.Item2];
        //                            rng.Formula = calcolo["SiglaInformazione"].Equals("CHECKINFO") ? DataBase.GetMessaggioCheck(risultato) : risultato;

        //                            if (calcolo["BackColor"] != DBNull.Value)
        //                                rng.Interior.ColorIndex = calcolo["BackColor"];
        //                            if (calcolo["ForeColor"] != DBNull.Value)
        //                                rng.Font.ColorIndex = calcolo["ForeColor"];

        //                            rng.ClearComments();

        //                            if (calcolo["Commento"] != DBNull.Value)
        //                                rng.AddComment(calcolo["Commento"]).Visible = false;
                                    
        //                            //TODO CHECK SE TUTTO FUNZIONA BENE...
        //                            //entitaInformazioni.RowFilter = "SiglaInformazione = '" + calcolo["SiglaInformazione"] + "'";
        //                            //if (entitaInformazioni.Count > 0 && entitaInformazioni[0]["SalvaDB"].Equals("1"))
        //                            //{
        //                                Handler.StoreEdit(ws, rng);
        //                            //}
        //                        }
        //                    }
        //                    watch.Stop();

        //                    if (calcolo["FineCalcolo"].Equals("1") || step == -1)
        //                        break;

        //                    if (calcolo["GoStep"] != DBNull.Value)
        //                        step = (int)calcolo["GoStep"];

        //                    if (step != 0)
        //                        i = calcoloRows.FindIndex(row => row["Step"].Equals(step));
        //                    else
        //                        i++;
        //                }
        //            }
        //        }
        //    }
        //}
        //private static object GetRisultatoCalcolo(object siglaEntita, DateTime data, int ora, DataRowView calcolo, Dictionary<object, object> entitaRiferimento, out int step)
        //{
        //    string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
        //    DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);
        //    Excel.Worksheet ws = _wb.Sheets.OfType<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name == nomeFoglio);

        //    string suffissoData = Date.GetSuffissoData(DataBase.DataAttiva, data);

        //    int ora1 = calcolo["OraInformazione1"] is DBNull ? ora : ora + (int)calcolo["OraInformazione1"];
        //    int ora2 = calcolo["OraInformazione2"] is DBNull ? ora : ora + (int)calcolo["OraInformazione2"];

        //    object siglaEntitaRif1 = calcolo["Riferimento1"] is DBNull ? (calcolo["SiglaEntita1"] is DBNull ? siglaEntita : calcolo["SiglaEntita1"]) : entitaRiferimento.FirstOrDefault(kv => kv.Value == calcolo["Riferimento1"]);
        //    object siglaEntitaRif2 = calcolo["Riferimento2"] is DBNull ? (calcolo["SiglaEntita2"] is DBNull ? siglaEntita : calcolo["SiglaEntita2"]) : entitaRiferimento.FirstOrDefault(kv => kv.Value == calcolo["Riferimento2"]);

        //    object valore1 = 0d;
        //    object valore2 = 0d;

        //    if (calcolo["SiglaInformazione1"] != DBNull.Value)
        //    {
        //        Tuple<int, int>[] riga = nomiDefiniti[DefinedNames.GetName(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData, "H" + ora1)];
        //        Tuple<int, int> cella = null;
        //        if (riga != null)
        //            cella = riga[0];

        //        switch (calcolo["SiglaInformazione1"].ToString())
        //        {
        //            case "UNIT_COMM":
        //                DataView entitaCommitment = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_COMMITMENT].DefaultView;
        //                entitaCommitment.RowFilter = "SiglaCommitment = '" + ws.Cells[cella.Item1, cella.Item2].Value + "'";
        //                valore1 = entitaCommitment.Count > 0 ? entitaCommitment[0]["IdEntitaCommitment"] : null;

        //                break;
        //            case "DISPONIBILITA":
        //                if (ws.Cells[cella.Item1, cella.Item2].Value == "OFF")
        //                    valore1 = 0d;
        //                else
        //                    valore1 = 1d;

        //                break;
        //            case "CHECKINFO":
        //                if (ws.Cells[cella.Item1, cella.Item2].Value == "OK")
        //                    valore1 = 1d;
        //                else
        //                    valore1 = 2d;
        //                break;
        //            default:
        //                if (cella != null)
        //                    valore1 = ws.Cells[cella.Item1, cella.Item2].Value ?? 0d;
        //                break;
        //        }
        //    }
        //    else if (calcolo["IdProprieta"] != DBNull.Value)
        //    {
        //        DataView entitaProprieta = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
        //        entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' IdProprieta = " + calcolo["IdProprieta"];

        //        if (entitaProprieta.Count > 0)
        //            valore1 = entitaProprieta[0]["Valore"];
        //    }
        //    else if (calcolo["IdParametroD"] != DBNull.Value)
        //    {
        //        DataView entitaParametro = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PARAMETRO_D].DefaultView;
        //        entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' IdParametro = " + calcolo["IdParametroD"];

        //        if (entitaParametro.Count > 0)
        //            valore1 = entitaParametro[0]["Valore"].ToString().Replace('.', ',');
        //    }
        //    else if (calcolo["IdParametroH"] != DBNull.Value)
        //    {
        //        DataView entitaParametro = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PARAMETRO_H].DefaultView;
        //        entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' IdParametro = " + calcolo["IdParametroH"];

        //        if (entitaParametro.Count > 0)
        //            valore1 = entitaParametro[0]["Valore"].ToString().Replace('.', ',');
        //    }
        //    else if (calcolo["Valore"] != DBNull.Value)
        //    {
        //        valore1 = Convert.ToDouble(calcolo["Valore"]);
        //    }

        //    if (calcolo["SiglaInformazione2"] != DBNull.Value)
        //    {
        //        Tuple<int, int>[] riga = nomiDefiniti[DefinedNames.GetName(siglaEntitaRif1, calcolo["SiglaInformazione2"], suffissoData, "H" + ora1)];
        //        Tuple<int, int> cella = null;
        //        if (riga != null)
        //            cella = riga[0];

        //        switch (calcolo["SiglaInformazione2"].ToString())
        //        {
        //            case "UNIT_COMM":
        //                DataView entitaCommitment = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_COMMITMENT].DefaultView;
        //                entitaCommitment.RowFilter = "SiglaCommitment = '" + ws.Cells[cella.Item1, cella.Item2].Value + "'";
        //                valore2 = entitaCommitment.Count > 0 ? entitaCommitment[0] : null;

        //                break;
        //            case "DISPONIBILITA":
        //                if (ws.Cells[cella.Item1, cella.Item2].Value == "OFF")
        //                    valore2 = 0d;
        //                else
        //                    valore2 = 1d;

        //                break;
        //            case "CHECKINFO":
        //                if (ws.Cells[cella.Item1, cella.Item2].Value == "OK")
        //                    valore2 = 1d;
        //                else
        //                    valore2 = 2d;
        //                break;
        //            default:
        //                if (cella != null)
        //                    valore2 = ws.Cells[cella.Item1, cella.Item2].Value ?? 0d;
        //                else
        //                    valore2 = 0d;
        //                break;
        //        }
        //    }

        //    double retVal = 0d;

        //    valore1 = valore1 ?? 0d;
        //    valore2 = valore2 ?? 0d;

        //    if (calcolo["Funzione"] is DBNull && calcolo["Operazione"] is DBNull && calcolo["Condizione"] is DBNull)
        //    {
        //        step = 0;
        //        if (Convert.ToDouble(valore1) == 0d)
        //            return valore2;

        //        return valore1;
        //    }
        //    else if (calcolo["Funzione"] != DBNull.Value)
        //    {
        //        string func = calcolo["Funzione"].ToString().ToLowerInvariant();
        //        if (calcolo["SiglaInformazione2"] is DBNull)
        //        {
        //            if (func.Contains("abs"))
        //            {
        //                retVal = Math.Abs(Convert.ToDouble(valore1));
        //            }
        //            else if (func.Contains("floor"))
        //            {
        //                retVal = Math.Floor(Convert.ToDouble(valore1));
        //            }
        //            else if (func.Contains("round"))
        //            {
        //                int decimals = int.Parse(func.Replace("round", ""));
        //                retVal = Math.Round(Convert.ToDouble(valore1), decimals);
        //            }
        //            else if (func.Contains("power"))
        //            {
        //                int exp = int.Parse(Regex.Match(func, @"\d*").Value);
        //                retVal = Math.Pow(Convert.ToDouble(valore1), exp);
        //            }
        //            else if (func.Contains("sum"))
        //            {
        //                foreach (var kvp in entitaRiferimento)
        //                {
        //                    Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, "H" + ora1)][0];
        //                    retVal += ws.Cells[cella.Item1, cella.Item2].Value ?? 0d;
        //                }
        //            }
        //            else if (func.Contains("avg"))
        //            {
        //                foreach (var kvp in entitaRiferimento)
        //                {
        //                    Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, "H" + ora1)][0];
        //                    retVal += ws.Cells[cella.Item1, cella.Item2].Value ?? 0d;
        //                }
        //                retVal /= entitaRiferimento.Count;
        //            }
        //            else if (func.Contains("max_h"))
        //            {
        //                retVal = double.MinValue;
        //                Tuple<int, int>[] riga = nomiDefiniti[DefinedNames.GetName(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData)];
        //                object[,] tmpVal = ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Value;
        //                object[] values = tmpVal.Cast<object>().ToArray();
        //                for (int i = 0; i < Date.GetOreGiorno(data); i++)
        //                {
        //                    double val = (double)(values[i] ?? 0);
        //                    retVal = Math.Max(val, retVal);
        //                }
        //            }
        //            else if (func.Contains("min_h"))
        //            {
        //                retVal = double.MaxValue;
        //                Tuple<int, int>[] riga = nomiDefiniti[DefinedNames.GetName(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData)];
        //                object[,] tmpVal = ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Value;
        //                object[] values = tmpVal.Cast<object>().ToArray();
        //                for (int i = 0; i < Date.GetOreGiorno(data); i++)
        //                {
        //                    double val = (double)(values[i] ?? 0);
        //                    retVal = Math.Min(val, retVal);
        //                }
        //            }
        //            else if (func.Contains("max"))
        //            {
        //                retVal = double.MinValue;
        //                foreach (var kvp in entitaRiferimento)
        //                {
        //                    Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, "H" + ora1)][0];
        //                    retVal = Math.Max(ws.Cells[cella.Item1, cella.Item2].Value ?? 0, retVal);
        //                }
        //            }
        //            else if (func.Contains("min"))
        //            {
        //                retVal = double.MaxValue;
        //                foreach (var kvp in entitaRiferimento)
        //                {
        //                    Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, "H" + ora1)][0];
        //                    retVal = Math.Min(ws.Cells[cella.Item1, cella.Item2].Value ?? 0, retVal);
        //                }
        //            }
        //        }
        //        //caso in cui ci sia anche SiglaInformazione2
        //        else
        //        {
        //            if (func.Contains("max"))
        //            {
        //                retVal = Math.Max(Convert.ToDouble(valore1), Convert.ToDouble(valore2));
        //            }
        //            else if (func.Contains("min"))
        //            {
        //                retVal = Math.Min(Convert.ToDouble(valore1), Convert.ToDouble(valore2));
        //            }
        //        }
        //    }
        //    else if (calcolo["Operazione"] != DBNull.Value)
        //    {
        //        switch (calcolo["Operazione"].ToString())
        //        {
        //            case "+":
        //                retVal = Convert.ToDouble(valore1) + Convert.ToDouble(valore2);
        //                break;
        //            case "-":
        //                retVal = Convert.ToDouble(valore1) - Convert.ToDouble(valore2);
        //                break;
        //            case "*":
        //                retVal = Convert.ToDouble(valore1) * Convert.ToDouble(valore2);
        //                break;
        //            case "/":
        //                retVal = Convert.ToDouble(valore1) / Convert.ToDouble(valore2);
        //                break;
        //        }
        //    }
        //    else if (calcolo["Condizione"] != DBNull.Value)
        //    {
        //        bool res = false;
        //        switch (calcolo["Condizione"].ToString())
        //        {
        //            case ">":
        //                res = Convert.ToDouble(valore1) > Convert.ToDouble(valore2);
        //                break;
        //            case "<":
        //                res = Convert.ToDouble(valore1) < Convert.ToDouble(valore2);
        //                break;
        //            case ">=":
        //                res = Convert.ToDouble(valore1) >= Convert.ToDouble(valore2);
        //                break;
        //            case "<=":
        //                res = Convert.ToDouble(valore1) <= Convert.ToDouble(valore2);
        //                break;
        //            case "=":
        //                res = Convert.ToDouble(valore1) == Convert.ToDouble(valore2);
        //                break;
        //            case "<>":
        //                res = Convert.ToDouble(valore1) != Convert.ToDouble(valore2);
        //                break;
        //        }
        //        if (res)
        //            step = (int)calcolo["StepCondizioneVera"];
        //        else
        //            step = (int)calcolo["StepCondizioneFalsa"];

        //        return res;
        //    }

        //    step = 0;
        //    return retVal;
        //}
        public static void InsertLog(Core.DataBase.TipologiaLOG logType, string message)
        {
            Excel.Worksheet log = _wb.Sheets["Log"];
            bool prot = log.ProtectContents;
            if (prot) log.Unprotect();
            DataBase.InsertLog(logType, message);
            if (prot) log.Protect();
        }
        public static void AggiornaLabelStatoDB()
        {
            bool isProtected = _wb.Sheets["Main"].ProtectContents;
            
            if (isProtected)
                _wb.Sheets["Main"].Unprotect(Simboli.pwd);

            if (DataBase.OpenConnection())
            {
                Dictionary<Core.DataBase.NomiDB, ConnectionState> stato = DataBase.DB.StatoDB;
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
                _wb.Sheets["Main"].Protect(Simboli.pwd);
        }
        public static void DumpDataSet()
        {
            StringWriter strWriter = new StringWriter();
            XmlWriter xmlWriter = XmlWriter.Create(strWriter);
            Utility.DataBase.LocalDB.WriteXml(xmlWriter, XmlWriteMode.WriteSchema);

            XElement root = XElement.Parse(strWriter.ToString());
            XNamespace ns = Simboli.NameSpace;

            IEnumerable<XElement> log =
                from tables in root.Elements(ns + Utility.DataBase.Tab.LOG)
                select tables;

            log.Remove();

            string locDBXml = strWriter.ToString();
            Microsoft.Office.Core.CustomXMLPart part;

            try
            {
                _wb.CustomXMLParts[Simboli.NameSpace].Delete();
            }
            catch
            {
            }
            part = _wb.CustomXMLParts.Add();

            part.LoadXML(root.ToString(SaveOptions.DisableFormatting));
            //part.LoadXML(locDBXml);
        }

        #endregion
    }

    public class Utilities : Workbook 
    {
        #region Variabili

        private static System.Version _wbVersion;

        #endregion

        #region Proprietà

        public static System.Version WorkbookVersion { get { return _wbVersion; } }
        public static System.Version CoreVersion { get { return DataBase.DB.GetCurrentV(); } }
        public static System.Version BaseVersion { get { return Assembly.GetExecutingAssembly().GetName().Version; } }

        #endregion

        #region Metodi

        public static void InitLog() 
        {
            DataTable dtLog = DataBase.DB.Select(DataBase.SP.APPLICAZIONE_LOG);
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

                DataTable dtUtente = DataBase.DB.Select(DataBase.SP.UTENTE, new QryParams() { { "@CodUtenteWindows", Environment.UserName } });
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
        public static bool Init(string dbName, object appID, DateTime dataAttiva, Microsoft.Office.Tools.Excel.Workbook wb, System.Version wbVersion) 
        {
            Core.CryptHelper.CryptSection("connectionStrings", "appSettings");

            DataBase.InitNewDB(dbName);
            DataBase.DB.PropertyChanged += _db_StatoDBChanged;
            DataBase.InitNewLocalDB();
            _wb = wb;
            _wbVersion = wbVersion;

            Simboli.pwd = ConfigurationManager.AppSettings["pwd"];

            bool localDBNotPresent = false;
            try
            {
                Office.CustomXMLPart xmlPart = _wb.CustomXMLParts[Simboli.NameSpace];
                StringReader sr = new StringReader(xmlPart.XML);
                DataBase.LocalDB.ReadXml(sr);
            }
            catch
            {
                localDBNotPresent = true;
                DataBase.LocalDB.Namespace = Simboli.NameSpace;
                //DataBase.LocalDB.Prefix = DataBase.NAME;
            }

            if (DataBase.OpenConnection())
            {
                Struttura.AggiornaParametriApplicazione(appID);

                int usr = InitUser();
                DataBase.DB.SetParameters(dataAttiva.ToString("yyyyMMdd"), usr, int.Parse(appID.ToString()));

                InitLog();
                DataBase.DB.CloseConnection();

                return false;
            }
            else //Emergenza
            {
                if (localDBNotPresent)
                {
                    System.Windows.Forms.MessageBox.Show("Il foglio non è inizializzato e non c'è connessione ad DB... Impossibile procedere! L'applicazione verrà chiusa.", "ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                    _wb.Close();
                }

                DataBase.DB.SetParameters(dataAttiva.ToString("yyyyMMdd"), 0, 0);
                DataView applicazione = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE].DefaultView;
                Simboli.nomeApplicazione = applicazione[0]["DesApplicazione"].ToString();
                Struct.intervalloGiorni = applicazione[0]["IntervalloGiorniEntita"] is DBNull ? 0 : int.Parse(applicazione[0]["IntervalloGiorniEntita"].ToString());
                Struct.visualizzaRiepilogo = applicazione[0]["VisRiepilogo"] is DBNull ? true : applicazione[0]["VisRiepilogo"].Equals("1");

                return true;
            }
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

        #endregion
    }
}
