using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Iren.FrontOffice.Core;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Tools.Excel;
using System.Deployment.Application;
using System.Reflection;

namespace Iren.FrontOffice.Base
{
    public class CommonFunctions
    {
        #region Costanti

        public static string NAME = "LocalDB";

        public struct Tab
        {
            public const string UTENTE = "Utente",
                LOG = "Log",
                AZIONE = "Azione",
                CATEGORIA = "Categoria",
                AZIONECATEGORIA = "AzioneCategoria",
                CATEGORIAENTITA = "CategoriaEntita",
                ENTITAAZIONE = "EntitaAzione",
                ENTITAINFORMAZIONE = "EntitaInformazione",
                ENTITAAZIONEINFORMAZIONE = "EntitaAzioneInformazione",
                CALCOLO = "Calcolo",
                CALCOLOINFORMAZIONE = "CalcoloInformazione",
                ENTITACALCOLO = "EntitaCalcolo",
                ENTITAGRAFICO = "EntitaGrafico",
                ENTITAGRAFICOINFORMAZIONE = "EntitaGraficoInformazione",
                ENTITACOMMITMENT = "EntitaCommitment",
                ENTITARAMPA = "EntitaRampa",
                ENTITAASSETTO = "EntitaAssetto",
                ENTITAPROPRIETA = "EntitaProprieta",
                ENTITAINFORMAZIONEFORMATTAZIONE = "EntitaInformazioneFormattazione",
                TIPOLOGIACHECK = "TipologiaCheck",
                TIPOLOGIARAMPA = "TipologiaRampa",
                APPLICAZIONE = "Applicazione",
                NOMIDEFINITI = "DefinedNames",
                ENTITAPARAMETROD = "EntitaParametroD",
                ENTITAPARAMETROH = "EntitaParametroH";
        }

        public enum AppIDs
        {
            PROGRAMMAZIONE_IMPIANTI = 5,
            SISTEMA_COMANDI = 8
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
        //public static Workbook ThisWorkBook
        //{
        //    get
        //    {
        //        return _wb;
        //    }
        //}
        
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

        #endregion

        #region Metodi

        private static void ResetTable(string name)
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

        private static void InitLog()
        {            
            ResetTable(Tab.LOG);
            DataTable dtLog = _db.Select("spApplicazioneLog");
            dtLog.TableName = Tab.LOG;
            dtLog.PrimaryKey = new DataColumn[] 
            {
                dtLog.Columns["Utente"],
                dtLog.Columns["Data"],
                dtLog.Columns["Testo"]

            };
            _localDB.Tables.Add(dtLog);
            
            DataView dv = _localDB.Tables[Tab.LOG].DefaultView;
            dv.Sort = "Data DESC";
        }

        private static DataTable CaricaApplicazione(AppIDs idApplicazione)
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

        public static void Init(string dbName, AppIDs appID, DateTime dataAttiva, Workbook wb, System.Version wbVersion)
        {
            _db = new DataBase(dbName);
            _localDB = new DataSet(NAME);
            _wb = wb;
            _wbVersion = wbVersion;

            DataTable dt = CaricaApplicazione(appID);
            if (dt.Rows.Count == 0)
                throw new ApplicationNotFoundException("L'appID inserito non ha restituito risultati.");
            
            _namespace = "Iren.ToolsExcel." + dt.Rows[0]["SiglaApplicazione"];
            Simboli.nomeApplicazione = dt.Rows[0]["DesApplicazione"].ToString();

            _localDB.Namespace = _namespace;
            _localDB.Prefix = NAME;
            _localDB.Tables.Add(dt);

            try
            {
                Office.CustomXMLPart xmlPart = _wb.CustomXMLParts[_namespace];
                StringReader sr = new StringReader(xmlPart.XML);
                _localDB.ReadXml(sr);
            }
            catch
            {
            }

            int usr = InitUser();
            _db.setParameters(dataAttiva.ToString("yyyyMMdd"), usr, (int)appID);

            InitLog();
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

        public static string GetSuffissoOra(object dataOra)
        {
            string dtO = dataOra.ToString();
            if (dtO.Length != 10)
                return "";

            return GetSuffissoOra(int.Parse(dtO.Substring(dtO.Length - 2, 2)));
        }

        public static string GetSuffissoOra(int ora)
        {
            return "H" + ora;
        }

        public static void AggiornaStrutturaDati()
        {
            CreaTabellaNomi();
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

        public static void Close()
        {
            StringWriter sw = new StringWriter();
            _localDB.WriteXml(sw);
            string locDBXml = sw.ToString();
            try
            {
                _wb.CustomXMLParts[_namespace].Delete();
            }
            catch
            {
            }
            _wb.CustomXMLParts.Add(locDBXml);
        }

        public static string GetName(params object[] parts)
        {
            string o = "";
            bool first = true;
            foreach (object part in parts)
            {
                o += (!first && part != "" ? Simboli.UNION : "") + part;
                first = false;
            }
            return o;
        }

        public static void InsertLog(DataBase.TipologiaLOG logType, string message)
        {
            _db.InsertLog(logType, message);
            DataTable dt = _db.Select("spApplicazioneLog");
            dt.TableName = Tab.LOG;
            _localDB.Merge(dt);
        }

        public void AggiornaFormule(Excel.Worksheet ws)
        {
            ws.Application.CalculateFull();
        }

        #endregion
    }
}
