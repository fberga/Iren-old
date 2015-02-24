﻿using Iren.FrontOffice.Core;
using Iren.FrontOffice.UserConfig;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Iren.FrontOffice.Base
{
    public class CommonFunctions
    {
        #region Costanti

        public static string NAME = "LocalDB";

        public struct Tab
        {
            public const string APPLICAZIONE = "Applicazione",
                AZIONE = "Azione",
                AZIONECATEGORIA = "AzioneCategoria",
                CALCOLO = "Calcolo",
                CALCOLOINFORMAZIONE = "CalcoloInformazione",
                CATEGORIA = "Categoria",
                CATEGORIAENTITA = "CategoriaEntita",
                ENTITAASSETTO = "EntitaAssetto",
                ENTITAAZIONE = "EntitaAzione",
                ENTITAAZIONEINFORMAZIONE = "EntitaAzioneInformazione",
                ENTITACALCOLO = "EntitaCalcolo",
                ENTITACOMMITMENT = "EntitaCommitment",
                ENTITAGRAFICO = "EntitaGrafico",
                ENTITAGRAFICOINFORMAZIONE = "EntitaGraficoInformazione",
                ENTITAINFORMAZIONE = "EntitaInformazione",
                ENTITAINFORMAZIONEFORMATTAZIONE = "EntitaInformazioneFormattazione",
                ENTITAPARAMETROD = "EntitaParametroD",
                ENTITAPARAMETROH = "EntitaParametroH",
                ENTITAPROPRIETA = "EntitaProprieta",
                ENTITARAMPA = "EntitaRampa",
                LOG = "Log",
                MODIFICA = "Modifica",
                NOMIDEFINITI = "DefinedNames",
                TIPOLOGIACHECK = "TipologiaCheck",
                TIPOLOGIARAMPA = "TipologiaRampa",
                UTENTE = "Utente";
        }

        #endregion

        #region Variabili

        private static string _namespace;
        private static DataSet _localDB = null;
        private static DataBase _db = null;
        private static Workbook _wb;
        private static System.Version _wbVersion;        

        #endregion

        #region Proprietà

        public static DataSet LocalDB 
        {
            get 
            {
                return _localDB;
            }
        }
        public static DataBase DB 
        {
            get 
            {
                return _db;
            } 
        }
        public static string NameSpace
        {
            get 
            { 
                return _namespace;
            }
        }
        public static System.Version CoreVersion
        {
            get { return _db.GetCurrentV(); }
        }
        public static System.Version BaseVersion
        {
            get { return Assembly.GetExecutingAssembly().GetName().Version; }
        }
        public static System.Version WorkbookVersion
        {
            get { return _wbVersion; }
        }
        public static Workbook WB { get { return _wb; } }

        #endregion

        #region Metodi

        public static void ResetTable(string name)
        {
            if (_localDB.Tables.Contains(name))
            {
                _localDB.Tables.Remove(name);
            }
        }

        private static int InitUser()
        {
            ResetTable(Tab.UTENTE);

            DataTable dtUtente = _db.Select("spUtente", new QryParams() { { "@CodUtenteWindows", Environment.UserName } });
            dtUtente.TableName = Tab.UTENTE;

            if (dtUtente.Rows.Count == 0)
            {
                DataRow r = dtUtente.NewRow();
                r["IdUtente"] = 0;
                r["Nome"] = "NON CONFIGURATO";
                dtUtente.Rows.Add(r);
            }
            _localDB.Tables.Add(dtUtente);

            return int.Parse(""+dtUtente.Rows[0]["IdUtente"]);
        }

        private static DataTable CaricaApplicazione(object idApplicazione)
        {
            string name = Tab.APPLICAZIONE;
            ResetTable(name);
            QryParams parameters = new QryParams() 
            {
                {"@IdApplicazione", idApplicazione},

            };
            DataTable dt = _db.Select(DataBase.StoredProcedure.APPLICAZIONE, parameters);
            dt.TableName = name;
            return dt;
        }

        public static void RefreshDate(DateTime dataAttiva)
        {
            _db.ChangeDate(dataAttiva.ToString("yyyyMMdd"));
        }

        public static void SwitchEnvironment(string ambiente)
        {
            RefreshAppSettings("DB", ambiente);
            _db = new DataBase(ambiente);
        }

        public static void InitLog()
        {
            DataTable dtLog = CommonFunctions.DB.Select("spApplicazioneLog");
            dtLog.TableName = CommonFunctions.Tab.LOG;
            CommonFunctions.LocalDB.Tables.Add(dtLog);

            DataView dv = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.LOG].DefaultView;
            dv.Sort = "Data DESC";

            CommonFunctions.DB.CloseConnection();
        }

        public static void Init(string dbName, object appID, DateTime dataAttiva, Workbook wb, System.Version wbVersion)
        {
            DataBase.CryptSection();
            _db = new DataBase(dbName);
            _localDB = new DataSet(NAME);
            _wb = wb;
            _wbVersion = wbVersion;

            if (_db.OpenConnection())
            {
                DataTable dt = CaricaApplicazione(appID);
                if (dt.Rows.Count == 0)
                    throw new ApplicationNotFoundException("L'appID inserito non ha restituito risultati.");

                _namespace = "Iren.ToolsExcel." + dt.Rows[0]["SiglaApplicazione"];
                Simboli.nomeApplicazione = dt.Rows[0]["DesApplicazione"].ToString();
                Simboli.intervalloGiorni = (dt.Rows[0]["IntervalloGiorni"] is DBNull ? 0 : (int)dt.Rows[0]["IntervalloGiorni"]);
                Simboli.pwd = ConfigurationManager.AppSettings["pwd"];

                try
                {
                    Office.CustomXMLPart xmlPart = _wb.CustomXMLParts[_namespace];
                    StringReader sr = new StringReader(xmlPart.XML);
                    _localDB.ReadXml(sr);
                    ResetTable(Tab.APPLICAZIONE);
                }
                catch
                {
                    _localDB.Namespace = _namespace;
                    _localDB.Prefix = NAME;
                }

                _localDB.Tables.Add(dt);

                int usr = InitUser();
                _db.SetParameters(dataAttiva.ToString("yyyyMMdd"), usr, int.Parse(appID.ToString()));

                InitLog();

                _db.CloseConnection();
            }
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
        
        public static string GetSuffissoOra(object dataOra)
        {
            string dtO = dataOra.ToString();
            if (dtO.Length != 10)
                return "";

            return GetSuffissoOra(int.Parse(dtO.Substring(dtO.Length - 2, 2)));
        }

        public static string GetDataFromSuffisso(object data, object ora = null)
        {
            int giorno = int.Parse(Regex.Match(data.ToString(), @"\d+").Value);
            DateTime outDate = DataBase.DataAttiva.AddDays(giorno - 1);

            ora = ora ?? "0";
            int outOra = int.Parse(Regex.Match(ora.ToString(), @"\d+").Value);

            return outDate.ToString("yyyyMMdd") + (outOra != 0 ? outOra.ToString("D2") : "");

        }

        public static string GetSuffissoOra(int ora)
        {
            return "H" + ora;
        }

        public static void ConvertiParametriInformazioni()
        {
            _db.Select("spApplicazioneInit");
        }

        public static void AggiornaStrutturaDati()
        {
            CreaTabellaNomi();
            CreaTabellaModifica();
            CaricaAzioni();
            CaricaCategorie();
            CaricaAzioneCategoria();
            CaricaCategoriaEntita();
            CaricaEntitaAzione();
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
                string name = Tab.NOMIDEFINITI;
                ResetTable(name);
                DataTable dt = DefinedNames.GetDefaultTable(name);
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

                dt.PrimaryKey = new DataColumn[] { dt.Columns["Entita"], dt.Columns["Informazione"], dt.Columns["Data"] };

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
                    {"@SiglaAzione", DataBase.ALL},
                    {"@Operativa", DataBase.ALL},
                    {"@Visibile", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.AZIONE, parameters);
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
                    {"@SiglaCategoria", DataBase.ALL},
                    {"@Operativa", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.CATEGORIA, parameters);
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
                string name = Tab.AZIONECATEGORIA;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaAzione", DataBase.ALL},
                    {"@SiglaCategoria", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.AZIONECATEGORIA, parameters);
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
                string name = Tab.CATEGORIAENTITA;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaCategoria", DataBase.ALL},
                    {"@SiglaEntita", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.CATEGORIAENTITA, parameters);
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
                string name = Tab.ENTITAAZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", DataBase.ALL},
                    {"@SiglaAzione", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.ENTITAAZIONE, parameters);
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
                string name = Tab.ENTITAINFORMAZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", DataBase.ALL},
                    {"@SiglaInformazione", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.ENTITAINFORMAZIONE, parameters);
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
                string name = Tab.ENTITAAZIONEINFORMAZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", DataBase.ALL},
                    {"@SiglaAzione", DataBase.ALL},
                    {"@SiglaInformazione", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.ENTITAAZIONEINFORMAZIONE, parameters);
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
                    {"@SiglaCalcolo", DataBase.ALL},
                    {"@IdTipologiaCalcolo", 0}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.CALCOLO, parameters);
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
                string name = Tab.CALCOLOINFORMAZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaCalcolo", DataBase.ALL},
                    {"@SiglaInformazione", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.CALCOLOINFORMAZIONE, parameters);
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
                string name = Tab.ENTITACALCOLO;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", DataBase.ALL},
                    {"@SiglaCalcolo", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.ENTITACALCOLO, parameters);
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
                string name = Tab.ENTITAGRAFICO;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", DataBase.ALL},
                    {"@SiglaGrafico", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.ENTITAGRAFICO, parameters);
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
                string name = Tab.ENTITAGRAFICOINFORMAZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", DataBase.ALL},
                    {"@SiglaGrafico", DataBase.ALL},
                    {"@SiglaInformazione", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.ENTITAGRAFICOINFORMAZIONE, parameters);
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
                string name = Tab.ENTITACOMMITMENT;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.ENTITACOMMITMENT, parameters);
                dt.TableName = name;
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
                string name = Tab.ENTITARAMPA;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.ENTITARAMPA, parameters);
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
                string name = Tab.ENTITAASSETTO;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.ENTITAASSETTO, parameters);
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
                string name = Tab.ENTITAPROPRIETA;
                ResetTable(name);
                QryParams parameters = new QryParams();
                DataTable dt = _db.Select(DataBase.StoredProcedure.ENTITAPROPRIETA, parameters);
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
                string name = Tab.ENTITAINFORMAZIONEFORMATTAZIONE;
                ResetTable(name);
                QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", DataBase.ALL},
                    {"@SiglaInformazione", DataBase.ALL}
                };
                DataTable dt = _db.Select(DataBase.StoredProcedure.ENTITAINFORMAZIONEFORMATTAZIONE, parameters);
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
                string name = Tab.TIPOLOGIACHECK;
                ResetTable(name);
                QryParams parameters = new QryParams();
                DataTable dt = _db.Select(DataBase.StoredProcedure.TIPOLOGIACHECK, parameters);
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
                string name = Tab.TIPOLOGIARAMPA;
                ResetTable(name);
                QryParams parameters = new QryParams();
                DataTable dt = _db.Select(DataBase.StoredProcedure.TIPOLOGIARAMPA, parameters);
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
                string name = Tab.ENTITAPARAMETROD;
                ResetTable(name);
                DataTable dt = _db.Select(DataBase.StoredProcedure.ENTITAPARAMETROD);
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
                string name = Tab.ENTITAPARAMETROH;
                ResetTable(name);
                DataTable dt = _db.Select(DataBase.StoredProcedure.ENTITAPARAMETROH);
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

        public static void DumpDataSet()
        {
            StringWriter strWriter = new StringWriter();
            XmlWriter xmlWriter = XmlWriter.Create(strWriter);
            _localDB.Tables.Remove(Tab.LOG);
            _localDB.WriteXml(xmlWriter);
            string locDBXml = strWriter.ToString();
            Microsoft.Office.Core.CustomXMLPart part;
            try
            {
                part = _wb.CustomXMLParts[_namespace];
            }
            catch
            {
                part = _wb.CustomXMLParts.Add();
            }

            part.LoadXML(locDBXml);
        }

        public static void InsertLog(DataBase.TipologiaLOG logType, string message)
        {
            if (_db.OpenConnection())
            {
                _wb.Sheets["Log"].Unprotect();
                _db.InsertLog(logType, message);
                //_db.CloseConnection();
                DataTable dt = _db.Select("spApplicazioneLog");
                dt.TableName = Tab.LOG;
                _localDB.Merge(dt);
                _wb.Sheets["Log"].Protect();
            }

        }

        public static void InsertApplicazioneRiepilogo(object siglaEntita, object siglaAzione, DateTime? dataRif = null, bool presente = true)
        {
            dataRif = dataRif ?? DataBase.DataAttiva;
            try
            {
                _db.OpenConnection();
                QryParams parameters = new QryParams() {
                    {"@SiglaEntita", siglaEntita},
                    {"@SiglaAzione", siglaAzione},
                    {"@Data", dataRif.Value.ToString("yyyyMMdd")},
                    {"@Presente", presente ? "1" : "0"}
                };
                _db.Insert("spInsertApplicazioneRiepilogo", parameters);
            }
            catch (Exception e)
            {
                //TODO riabilitare log
                //InsertLog(DataBase.TipologiaLOG.LogErrore, "InsertApplicazioneRiepilogo ["+ dataRif ?? DataBase.DataAttiva +", " + siglaEntita + ", " + siglaAzione + "]: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public static void AggiornaFormule(Excel.Worksheet ws)
        {
            ws.Application.CalculateFull();
        }

        public static void RefreshAppSettings(string key, string value)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings[key].Value = value;
            config.Save(ConfigurationSaveMode.Minimal);
            ConfigurationManager.RefreshSection("appSettings");
        }

        public static bool CaricaAzioneInformazione(object siglaEntita, object siglaAzione, object azionePadre, DateTime? dataRif = null, object parametro = null)
        {
            try
            {
                DataView azioni = _localDB.Tables[Tab.AZIONE].DefaultView;
                azioni.RowFilter = "SiglaAzione = '" + siglaAzione + "'";

                if (dataRif == null)
                    dataRif = DataBase.DataAttiva;

                bool procedi = true;
                if (azioni[0]["Visibile"].Equals("1"))
                {
                    procedi = false;
                    DataView entitaAzioni = _localDB.Tables[Tab.ENTITAAZIONE].DefaultView;
                    entitaAzioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";
                    if (entitaAzioni.Count > 0)
                    {
                        procedi = true;
                    }
                }

                if (procedi)
                {
                    AzzeraInformazione(siglaEntita, siglaAzione, dataRif);

                    if (_db.OpenConnection())
                    {
                        if (azionePadre.Equals("GENERA"))
                        {
                            ElaborazioneInformazione(siglaEntita, dataRif.Value, (siglaAzione.Equals("G_MP_MGP") ? 5 : 7));
                            if (azioni[0]["Visibile"].Equals("1"))
                                InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, dataRif);
                        }
                        else
                        {
                            DataView azioneInformazione = _db.Select("spCaricaAzioneInformazione", "@SiglaEntita=" + siglaEntita + ";@SiglaAzione=" + siglaAzione + ";@Parametro=" + parametro + ";@Data=" + dataRif.Value.ToString("yyyyMMdd")).DefaultView;
                            if (azioneInformazione.Count == 0)
                            {
                                if (azioni[0]["Visibile"].Equals("1"))
                                    InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, dataRif, false);
                            }
                            else
                            {
                                ScriviInformazione(siglaEntita, azioneInformazione);

                                if (azioni[0]["Visibile"].Equals("1"))
                                    InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, dataRif);
                            }
                        }
                    }
                    else
                    {
                        if (azionePadre.Equals("GENERA"))
                            ElaborazioneInformazione(siglaEntita, dataRif.Value, (siglaAzione.Equals("G_MP_MGP") ? 5 : 7));
                    }

                }
                return true;
            }
            catch (Exception e)
            {
                //TODO riabilitare log!!
                //InsertLog(DataBase.TipologiaLOG.LogErrore, "modProgram CaricaAzioneInformazione [" + siglaEntita + ", " + siglaAzione + "]: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        public static void AzzeraInformazione(object siglaEntita, object siglaAzione, DateTime? dataRif = null, object valore = null)
        {
            string foglio = DefinedNames.GetSheetName(siglaEntita);

            DefinedNames nomiDefiniti = new DefinedNames(foglio);
            Excel.Worksheet ws = _wb.Sheets.OfType<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name == foglio);

            if (dataRif == null)
                dataRif = DataBase.DataAttiva;

            string suffissoData = GetSuffissoData(DataBase.DataAttiva, dataRif.Value);

            DataView entitaAzioniInformazioni = _localDB.Tables[Tab.ENTITAAZIONEINFORMAZIONE].DefaultView;

            //TODO controllare perché Domenico passa un entitaRif a true/false in questo filtro
            entitaAzioniInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";

            foreach (DataRowView entitaAzioneInformazione in entitaAzioniInformazioni)
            {
                if (entitaAzioneInformazione["FormulaInCella"].Equals("0"))
                {
                    object entita = entitaAzioneInformazione["SiglaEntitaRif"] is DBNull ? entitaAzioneInformazione["SiglaEntita"] : entitaAzioneInformazione["SiglaEntitaRif"];
                    Tuple<int, int>[] riga = new Tuple<int, int>[0];

                    if (entitaAzioneInformazione["Selezione"].Equals(0))
                        riga = nomiDefiniti[DefinedNames.GetName(entita, entitaAzioneInformazione["SiglaInformazione"], suffissoData)];
                    else
                        riga = nomiDefiniti[DefinedNames.GetName(entita, "SEL", entitaAzioneInformazione["Selezione"], suffissoData)];

                    Excel.Range rng = ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]];
                    rng.Value = valore;
                    rng.Interior.ColorIndex = entitaAzioneInformazione["BackColor"];
                    rng.Font.ColorIndex = entitaAzioneInformazione["ForeColor"];
                    rng.ClearComments();
                }
            }
        }

        public static void ScriviInformazione(object siglaEntita, DataView azioneInformazione)
        {
            string foglio = DefinedNames.GetSheetName(siglaEntita);

            DefinedNames nomiDefiniti = new DefinedNames(foglio);
            Excel.Worksheet ws = _wb.Sheets.OfType<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name == foglio);

            foreach (DataRowView azione in azioneInformazione)
            {
                string suffissoData;
                if (azione["SiglaEntita"].Equals("UP_BUS") && azione["SiglaInformazione"].Equals("VOL_INVASO"))
                    suffissoData = DefinedNames.GetName("DATA0", "H24");
                else 
                    suffissoData = DefinedNames.GetName(GetSuffissoData(DataBase.DataAttiva, azione["Data"]), GetSuffissoOra(azione["Data"]));

                Tuple<int, int>[] celle = nomiDefiniti[DefinedNames.GetName(azione["SiglaEntita"], azione["SiglaInformazione"], suffissoData)];
                if (celle != null)
                {
                    Excel.Range rng = ws.Cells[celle[0].Item1, celle[0].Item2];
                    rng.Value = azione["Valore"];
                    if (azione["BackColor"] != DBNull.Value)
                        rng.Interior.ColorIndex = azione["BackColor"];
                    if (azione["BackColor"] != DBNull.Value)
                        rng.Font.ColorIndex = azione["ForeColor"];
                    
                    rng.ClearComments();
                    
                    if (azione["Commento"] != DBNull.Value)
                        rng.AddComment(azione["Commento"]).Visible = false;
                }
            }
        }

        private static void ElaborazioneInformazione(object siglaEntita, DateTime data, int tipologiaCalcolo, int oraInizio = 0, int oraFine = 0)
        {
            Dictionary<object, object> entitaRiferimento = new Dictionary<object, object>();
            List<int> oreDaCalcolare = new List<int>();
            

            string suffissoData = GetSuffissoData(DataBase.DataAttiva, data);
            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            Excel.Worksheet ws = _wb.Sheets.OfType<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name == nomeFoglio);
            DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);

            if (oraInizio == 0)
            {
                oraInizio++;
                oraFine = GetOreGiorno(data);
            }
            DataView categoriaEntita = _localDB.Tables[Tab.CATEGORIAENTITA].DefaultView;
            DataView entitaInformazioni = _localDB.Tables[Tab.ENTITAINFORMAZIONE].DefaultView;
            categoriaEntita.RowFilter = "Gerarchia = '" + siglaEntita + "'";
            foreach (DataRowView entita in categoriaEntita)
                entitaRiferimento.Add(entita["SiglaEntita"].ToString(), entita["Riferimento"]);

            if (entitaRiferimento.Count == 0)
                entitaRiferimento.Add(siglaEntita, 1);


            if (tipologiaCalcolo == 1 || tipologiaCalcolo == 5 && DefinedNames.IsDefined(nomeFoglio, DefinedNames.GetName(siglaEntita, "UNIT_COMM")))
            {
                DataView entitaCommitment = _localDB.Tables[Tab.ENTITACOMMITMENT].DefaultView;

                Tuple<int, int> primaCella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "UNIT_COMM", suffissoData, "H" + oraInizio)][0];
                Tuple<int, int> ultimaCella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "UNIT_COMM", suffissoData, "H" + oraFine)][0];
                object[,] tmpVal = ws.Range[ws.Cells[primaCella.Item1, primaCella.Item2], ws.Cells[ultimaCella.Item1, ultimaCella.Item2]].Value;
                object[] values = tmpVal.Cast<object>().ToArray();
                for (int i = oraInizio; i < oraFine; i++)
                {
                    entitaCommitment.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaCommitment = '" + values[i - oraInizio] + "' AND AbilitaOffertaMGP = '1'";
                    if (entitaCommitment.Count > 0)
                        oreDaCalcolare.Add(i);
                }
            }
            else
            {
                for (int i = oraInizio; i < oraFine; i++)
                    oreDaCalcolare.Add(i);
            }

            if (oreDaCalcolare.Count > 0)
            {
                if (tipologiaCalcolo == 3)
                {
                    foreach (int ora in oreDaCalcolare)
                    {
                        Tuple<int,int> cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "CHECKINFO", suffissoData, "H" + ora)][0];
                        ws.Range[cella.Item1, cella.Item2].Value = null;
                    }
                }

                DataView entitaCalcoli = _localDB.Tables[Tab.ENTITACALCOLO].DefaultView;
                entitaCalcoli.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdTipologiaCalcolo = " + tipologiaCalcolo;

                DataView calcoloInformazioni = _localDB.Tables[Tab.CALCOLOINFORMAZIONE].DefaultView;

                foreach (DataRowView entitaCalcolo in entitaCalcoli)
                {
                    calcoloInformazioni.RowFilter = "SiglaCalcolo = '" + entitaCalcolo["SiglaCalcolo"] + "'";
                    DataView tmp = calcoloInformazioni.ToTable(true, "SiglaEntitaRif", "SiglaInformazione").DefaultView;

                    foreach (int ora in oreDaCalcolare)
                    {
                        //foreach (DataRowView calcoloInfo in tmp)
                        //{
                        //    if (!calcoloInfo["SiglaInformazione"].Equals("CHECKINFO"))
                        //    {
                        //        object siglaEntita2 = calcoloInfo["SiglaEntitaRif"] is DBNull ? siglaEntita : calcoloInfo["SiglaEntitaRif"];
                        //        if (DefinedNames.IsDefined(nomeFoglio, GetName(siglaEntita2, calcoloInfo["SiglaInformazione"], suffissoData, "H" + ora)))
                        //        {
                        //            Tuple<int, int> cella = nomiDefiniti[GetName(siglaEntita2, calcoloInfo["SiglaInformazione"], suffissoData, "H" + ora)][0];
                        //            ws.Cells[cella.Item1, cella.Item2].Value = null;
                        //        }
                        //    }
                        //}

                        List<DataRowView> calcoloRows = calcoloInformazioni.Cast<DataRowView>().ToList();
                        int i = 0;

                        while (i < calcoloRows.Count)
                        {
                            DataRowView calcolo = calcoloRows[i];

                            if (calcolo["OraInizio"] != DBNull.Value)
                                if (ora < oraInizio && ora > oraFine)
                                {
                                    i++;
                                    continue;
                                }
                                    

                            if (calcolo["OraFine"].Equals("0"))
                            {
                                if (ora != GetOreGiorno(data))
                                    break;
                                i++;
                                continue;
                            }

                            int step;
                            Stopwatch watch = Stopwatch.StartNew();
                            object risultato = GetRisultatoCalcolo(siglaEntita, data, ora, calcolo, entitaRiferimento, out step);
                            watch.Stop();
                            watch = Stopwatch.StartNew();
                            if (step == 0)
                            {
                                if (nomiDefiniti.IsDefined(DefinedNames.GetName(siglaEntita, calcolo["SiglaInformazione"], suffissoData, "H" + ora)))
                                {
                                    Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, calcolo["SiglaInformazione"], suffissoData, "H" + ora)][0];
                                
                                    Excel.Range rng = ws.Cells[cella.Item1, cella.Item2];
                                    rng.Formula = calcolo["SiglaInformazione"].Equals("CHECKINFO") ? GetMessaggioCheck(risultato) : risultato;

                                    if (calcolo["BackColor"] != DBNull.Value)
                                        rng.Interior.ColorIndex = calcolo["BackColor"];
                                    if (calcolo["ForeColor"] != DBNull.Value)
                                        rng.Font.ColorIndex = calcolo["ForeColor"];

                                    rng.ClearComments();

                                    if (calcolo["Commento"] != DBNull.Value)
                                        rng.AddComment(calcolo["Commento"]).Visible = false;

                                    entitaInformazioni.RowFilter = "SiglaInformazione = '" + calcolo["SiglaInformazione"] + "'";
                                    if (entitaInformazioni.Count > 0 && entitaInformazioni[0]["SalvaDB"].Equals("1"))
                                    {
                                        BaseHandler.StoreEdit(ws, rng);
                                    }
                                }
                            }
                            watch.Stop();

                            if (calcolo["FineCalcolo"].Equals("1") || step == -1)
                                break;

                            if (calcolo["GoStep"] != DBNull.Value)
                                step = (int)calcolo["GoStep"];

                            if (step != 0) 
                                i = calcoloRows.FindIndex(row => row["Step"].Equals(step));
                            else
                                i++;
                        }
                    }
                }
            }
        }

        private static object GetRisultatoCalcolo(object siglaEntita, DateTime data, int ora, DataRowView calcolo, Dictionary<object, object> entitaRiferimento, out int step)
        {
            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);
            Excel.Worksheet ws = _wb.Sheets.OfType<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name == nomeFoglio);

            string suffissoData = GetSuffissoData(DataBase.DataAttiva, data);

            int ora1 = calcolo["OraInformazione1"] is DBNull ? ora : ora + (int)calcolo["OraInformazione1"];
            int ora2 = calcolo["OraInformazione2"] is DBNull ? ora : ora + (int)calcolo["OraInformazione2"];

            object siglaEntitaRif1 = calcolo["Riferimento1"] is DBNull ? (calcolo["SiglaEntita1"] is DBNull ? siglaEntita : calcolo["SiglaEntita1"]) : entitaRiferimento.FirstOrDefault(kv => kv.Value == calcolo["Riferimento1"]);
            object siglaEntitaRif2 = calcolo["Riferimento2"] is DBNull ? (calcolo["SiglaEntita2"] is DBNull ? siglaEntita : calcolo["SiglaEntita2"]) : entitaRiferimento.FirstOrDefault(kv => kv.Value == calcolo["Riferimento2"]);

            object valore1 = 0d;
            object valore2 = 0d;

            if (calcolo["SiglaInformazione1"] != DBNull.Value)
            {
                Tuple<int, int>[] riga = nomiDefiniti[DefinedNames.GetName(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData, "H" + ora1)];
                Tuple<int,int> cella = null;
                if(riga != null)
                    cella = riga[0];

                switch (calcolo["SiglaInformazione1"].ToString())
                {
                    case "UNIT_COMM":
                        DataView entitaCommitment = _localDB.Tables[Tab.ENTITACOMMITMENT].DefaultView;
                        entitaCommitment.RowFilter = "SiglaCommitment = '" + ws.Cells[cella.Item1, cella.Item2].Value + "'";
                        valore1 = entitaCommitment.Count > 0 ? entitaCommitment[0]["IdEntitaCommitment"] : null;
                        
                        break;
                    case "DISPONIBILITA":
                        if (ws.Cells[cella.Item1, cella.Item2].Value == "OFF")
                            valore1 = 0d;
                        else
                            valore1 = 1d;

                        break;
                    case "CHECKINFO":
                        if (ws.Cells[cella.Item1, cella.Item2].Value == "OK")
                            valore1 = 1d;
                        else
                            valore1 = 2d;
                        break;
                    default:
                        if (cella != null)
                            valore1 = ws.Cells[cella.Item1, cella.Item2].Value ?? 0d;
                        break;
                }
            }
            else if (calcolo["IdProprieta"] != DBNull.Value)
            {
                DataView entitaProprieta = _localDB.Tables[Tab.ENTITAPROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' IdProprieta = " + calcolo["IdProprieta"];

                if (entitaProprieta.Count > 0)
                    valore1 = entitaProprieta[0]["Valore"];
            }
            else if (calcolo["IdParametroD"] != DBNull.Value)
            {
                DataView entitaParametro = _localDB.Tables[Tab.ENTITAPARAMETROD].DefaultView;
                entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' IdParametro = " + calcolo["IdParametroD"];

                if (entitaParametro.Count > 0)
                    valore1 = entitaParametro[0]["Valore"].ToString().Replace('.', ',');
            }
            else if (calcolo["IdParametroH"] != DBNull.Value)
            {
                DataView entitaParametro = _localDB.Tables[Tab.ENTITAPARAMETROH].DefaultView;
                entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' IdParametro = " + calcolo["IdParametroH"];

                if (entitaParametro.Count > 0)
                    valore1 = entitaParametro[0]["Valore"].ToString().Replace('.', ',');
            }
            else if(calcolo["Valore"] != DBNull.Value)
            {
                valore1 = Convert.ToDouble(calcolo["Valore"]);
            }

            if (calcolo["SiglaInformazione2"] != DBNull.Value)
            {
                Tuple<int, int>[] riga = nomiDefiniti[DefinedNames.GetName(siglaEntitaRif1, calcolo["SiglaInformazione2"], suffissoData, "H" + ora1)];
                Tuple<int, int> cella = null;
                if (riga != null)
                    cella = riga[0];

                switch (calcolo["SiglaInformazione2"].ToString())
                {
                    case "UNIT_COMM":
                        DataView entitaCommitment = _localDB.Tables[Tab.ENTITACOMMITMENT].DefaultView;
                        entitaCommitment.RowFilter = "SiglaCommitment = '" + ws.Cells[cella.Item1, cella.Item2].Value + "'";
                        valore2 = entitaCommitment.Count > 0 ? entitaCommitment[0] : null;

                        break;
                    case "DISPONIBILITA":
                        if (ws.Cells[cella.Item1, cella.Item2].Value == "OFF")
                            valore2 = 0d;
                        else
                            valore2 = 1d;

                        break;
                    case "CHECKINFO":
                        if (ws.Cells[cella.Item1, cella.Item2].Value == "OK")
                            valore2 = 1d;
                        else
                            valore2 = 2d;
                        break;
                    default:
                        if (cella != null)
                            valore2 = ws.Cells[cella.Item1, cella.Item2].Value ?? 0d;
                        else
                            valore2 = 0d;
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
                    if(func.Contains("abs")) 
                    {
                        retVal = Math.Abs(Convert.ToDouble(valore1));
                    }
                    else if (func.Contains("floor"))
                    {
                        retVal = Math.Floor(Convert.ToDouble(valore1));
                    }
                    else if (func.Contains("round"))
                    {
                        int decimals = int.Parse(func.Replace("round",""));
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
                        {
                            Tuple<int,int> cella = nomiDefiniti[DefinedNames.GetName(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, "H" + ora1)][0];
                            retVal += ws.Cells[cella.Item1, cella.Item2].Value ?? 0d;
                        }
                    }
                    else if (func.Contains("avg"))
                    {
                        foreach (var kvp in entitaRiferimento)
                        {
                            Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, "H" + ora1)][0];
                            retVal += ws.Cells[cella.Item1, cella.Item2].Value ?? 0d;
                        }
                        retVal /= entitaRiferimento.Count;
                    }
                    else if (func.Contains("max_h"))
                    {
                        retVal = double.MinValue;
                        Tuple<int, int>[] riga = nomiDefiniti[DefinedNames.GetName(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData)];
                        object[,] tmpVal = ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Value;
                        object[] values = tmpVal.Cast<object>().ToArray();
                        for (int i = 0; i < GetOreGiorno(data); i++)
                        {
                            double val = (double)(values[i] ?? 0);
                            retVal = Math.Max(val, retVal);
                        }
                    }
                    else if (func.Contains("min_h"))
                    {
                        retVal = double.MaxValue;
                        Tuple<int, int>[] riga = nomiDefiniti[DefinedNames.GetName(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData)];
                        object[,] tmpVal = ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Value;
                        object[] values = tmpVal.Cast<object>().ToArray();
                        for (int i = 0; i < GetOreGiorno(data); i++)
                        {
                            double val = (double)(values[i] ?? 0);
                            retVal = Math.Min(val, retVal);
                        }
                    }
                    else if (func.Contains("max"))
                    {
                        retVal = double.MinValue;
                        foreach (var kvp in entitaRiferimento)
                        {
                            Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, "H" + ora1)][0];
                            retVal = Math.Max(ws.Cells[cella.Item1, cella.Item2].Value ?? 0, retVal);
                        }
                    }
                    else if (func.Contains("min"))
                    {
                        retVal = double.MaxValue;
                        foreach (var kvp in entitaRiferimento)
                        {
                            Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, "H" + ora1)][0];
                            retVal = Math.Min(ws.Cells[cella.Item1, cella.Item2].Value ?? 0, retVal);
                        }
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

        private static object GetMessaggioCheck(object id) 
        {
            DataView tipologiaCheck = _localDB.Tables[Tab.TIPOLOGIACHECK].DefaultView;
            tipologiaCheck.RowFilter = "IdTipologiaCheck = " + id;

            if(tipologiaCheck.Count > 0)
                return tipologiaCheck[0]["Messaggio"];

            return null;
        }

        public static void SalvaModificheDB()
        {
            DataTable modifiche = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.MODIFICA];

            DataTable dt = modifiche.Copy();
            dt.TableName = modifiche.TableName;
            dt.Namespace = "";

            if (dt.Rows.Count == 0)
                return;

            bool onLine = DB.OpenConnection();

            var path = Esporta.GetPath("pathExportModifiche");

            string cartellaRemota = Esporta.PreparePath(path.Value);
            string cartellaEmergenza = Esporta.PreparePath(path.Emergenza);
            string cartellaArchivio = Esporta.PreparePath(path.Archivio);

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
                        if(true)
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
        
        #endregion
    }
}
