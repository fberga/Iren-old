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
using System.Deployment.Application;
using System.Reflection;
using System.Configuration;

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

        public static void InitLog()
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

        public static void ChangeModificaDati(bool modifica)
        {
            Excel.Worksheet ws = _wb.Sheets["Main"];

            ws.Shapes.Item("lbModifica").TextFrame.Characters().Text = "Modifica dati: " + (modifica ? "SI" : "NO");
        }

        public static void SwitchEnvironment(string ambiente)
        {
            RefreshAppSettings("DB", ambiente);
            _db = new DataBase(ambiente);
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
            Simboli.intervalloGiorni = (dt.Rows[0]["IntervalloGiorni"] is DBNull ? 0 : (int)dt.Rows[0]["IntervalloGiorni"]);

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

        public static void ConvertiParametriInformazioni()
        {
            _db.Select("spApplicazioneInit");
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

        public static void RefreshAppSettings(string key, string value)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings[key].Value = value;
            config.Save(ConfigurationSaveMode.Minimal);
            ConfigurationManager.RefreshSection("appSettings");
        }

        public static void CaricaAzioneInformazione(object siglaEntita, object siglaAzione, object azionePadre, DateTime? dataRif = null, object parametro = null)
        {
            DataView azioni = _localDB.Tables[Tab.AZIONE].DefaultView;
            azioni.RowFilter = "SiglaAzione = '" + siglaAzione + "'";

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
                AzzeraInformazione(siglaEntita, siglaAzione);

                if (_db.StatoDB()[DataBase.NomiDB.SQLSERVER] == ConnectionState.Open)
                {
                    if (azionePadre.Equals("GENERA"))
                    {

                    }
                }

            }
        }

        public static void AzzeraInformazione(object siglaEntita, object siglaAzione, DateTime? dataRif = null, object valore = null)
        {
            string foglio = DefinedNames.GetSheetName(siglaEntita);

            DefinedNames nomiDefiniti = new DefinedNames(foglio);
            Excel.Worksheet ws = _wb.Sheets.OfType<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name == foglio);

            if (dataRif == null)
                dataRif = DataBase.Data;

            string suffissoData = GetSuffissoData(DataBase.Data, dataRif.Value);

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
                        riga = nomiDefiniti[GetName(entita, entitaAzioneInformazione["SiglaInformazione"])];
                    else
                        riga = nomiDefiniti[GetName(entita, "SEL", entitaAzioneInformazione["Selezione"])];

                    Excel.Range rng = ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]];
                    rng.Value = valore;
                    rng.Interior.ColorIndex = entitaAzioneInformazione["BackColor"];
                    rng.Font.ColorIndex = entitaAzioneInformazione["ForeColor"];
                    rng.ClearComments();
                }
            }
        }

        public static string R1C1toA1(int riga, int colonna)
        {
            string output = "";
            while (colonna > 0)
            {
                int lettera = colonna % 26;
                output = char.ConvertFromUtf32(lettera + 64) + output;
                colonna = colonna / 26;
            }
            output += riga;
            return output;
        }

        private static void ElaborazioneInformazione(string siglaEntita, DateTime data, int tipologiaCalcolo, int oraInizio = 0, int oraFine = 0)
        {
            Dictionary<object, object> entitaRiferimento = new Dictionary<object, object>();
            List<int> oreDaCalcolare = new List<int>();
            

            string suffissoData = GetSuffissoData(DataBase.Data, data);
            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            Excel.Worksheet ws = _wb.Sheets.OfType<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name == nomeFoglio);
            DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);

            if (oraInizio == 0)
            {
                oraInizio++;
                oraFine = GetOreGiorno(data);
            }
            DataView categoriaEntita = _localDB.Tables[Tab.CATEGORIAENTITA].DefaultView;
            categoriaEntita.RowFilter = "Gerarchia = '" + siglaEntita + "'";
            foreach (DataRowView entita in categoriaEntita)
                entitaRiferimento.Add(entita["SiglaEntita"].ToString(), entita["Riferimento"]);

            if (entitaRiferimento.Count == 0)
                entitaRiferimento.Add(siglaEntita, 1);


            if (tipologiaCalcolo == 1 || tipologiaCalcolo == 5 && DefinedNames.IsDefined(nomeFoglio, GetName(siglaEntita, "UNIT_COMM")))
            {
                DataView entitaCommitment = _localDB.Tables[Tab.ENTITACOMMITMENT].DefaultView;

                Tuple<int, int> primaCella = nomiDefiniti[GetName(siglaEntita, "UNIT_COMM", suffissoData, "H" + oraInizio)][0];
                Tuple<int, int> ultimaCella = nomiDefiniti[GetName(siglaEntita, "UNIT_COMM", suffissoData, "H" + oraFine)][0];
                object[,] values = ws.Range[ws.Cells[primaCella.Item1, primaCella.Item2], ws.Cells[ultimaCella.Item1, ultimaCella.Item2]].Value;

                for (int i = oraInizio; i < oraFine; i++)
                {
                    entitaCommitment.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaCommitment = '" + values[0, i - oraInizio] + "' AND AbilitaOfferta = '1'";
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
                        Tuple<int,int> cella = nomiDefiniti[GetName(siglaEntita, "CHECKINFO", suffissoData, "H" + ora)][0];
                        ws.Range[cella.Item1, cella.Item2].Value = null;
                    }
                }

                DataView entitaCalcoli = _localDB.Tables[Tab.ENTITACALCOLO].DefaultView;
                entitaCalcoli.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdTipologiaCalcolo = " + tipologiaCalcolo;

                DataView calcoloInformazioni = _localDB.Tables[Tab.CALCOLOINFORMAZIONE].DefaultView;


                foreach (DataRowView entitaCalcolo in entitaCalcoli)
                {
                    calcoloInformazioni.RowFilter = "SiglaCalcolo = '" + entitaCalcolo["SiglaCalcolo"] + "'";

                    foreach (int ora in oreDaCalcolare)
                    {
                        foreach (DataRowView calcoloInfo in calcoloInformazioni)
                        {
                            if (!entitaCalcolo["SiglaInformazione"].Equals("CHECKINFO"))
                            {
                                object siglaEntita2 = entitaCalcolo["SiglaEntitaRif"] is DBNull ? siglaEntita : entitaCalcolo["SiglaEntitaRif"];
                                Tuple<int, int> cella = nomiDefiniti[GetName(siglaEntita, "CHECKINFO", suffissoData, "H" + ora)][0];
                                ws.Cells[cella.Item1, cella.Item2].Value = null;
                            }
                        }

                        foreach (DataRowView calcoloInfo in calcoloInformazioni)
                        {
                            if (calcoloInfo["OraInizio"] != DBNull.Value)
                                if (ora < oraInizio && ora > oraFine)
                                    continue;

                            if (calcoloInfo["OraFine"].Equals("0"))
                            {
                                if (ora != GetOreGiorno(data))
                                    break;
                                continue;
                            }

                            //get risultato calcolo...


                        }
                    }
                }
            }
        }

        private static void GetRisultatoCalcolo(object siglaEntita, DateTime data, int ora, DataRowView calcolo, Dictionary<object, object> entitaRiferimento)
        {
            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);
            Excel.Worksheet ws = _wb.Sheets.OfType<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name == nomeFoglio);

            string suffissoData = GetSuffissoData(DataBase.Data, data);

            int ora1 = calcolo["OraInformazione1"] is DBNull ? ora : ora + (int)calcolo["OraInformazione1"];
            int ora2 = calcolo["OraInformazione2"] is DBNull ? ora : ora + (int)calcolo["OraInformazione2"];

            object siglaEntitaRif1 = calcolo["Riferimento1"] is DBNull ? (calcolo["SiglaEntita1"] is DBNull ? siglaEntita : calcolo["SiglaEntita1"]) : entitaRiferimento.FirstOrDefault(kv => kv.Value == calcolo["Riferimento1"]);
            object siglaEntitaRif2 = calcolo["Riferimento2"] is DBNull ? (calcolo["SiglaEntita2"] is DBNull ? siglaEntita : calcolo["SiglaEntita2"]) : entitaRiferimento.FirstOrDefault(kv => kv.Value == calcolo["Riferimento2"]);

            object valore1 = null;
            object valore2;

            if (calcolo["SiglaInformazione1"] != DBNull.Value)
            {
                Tuple<int, int>[] riga = nomiDefiniti[GetName(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData, "H" + ora1)];
                Tuple<int,int> cella = null;
                if(riga != null)
                    cella = riga[0];

                switch (calcolo["SiglaInformazione1"].ToString())
                {
                    case "UNIT_COMM":
                        DataView entitaCommitment = _localDB.Tables[Tab.ENTITACOMMITMENT].DefaultView;
                        entitaCommitment.RowFilter = "SiglaCommitment = '" + ws.Cells[cella.Item1, cella.Item2].Value + "'";
                        valore1 = entitaCommitment.Count > 0 ? entitaCommitment[0] : null;
                        
                        break;
                    case "DISPONIBILITA":
                        if (ws.Cells[cella.Item1, cella.Item2].Value == "OFF")
                            valore1 = 0;
                        else
                            valore1 = 1;

                        break;
                    case "CHECKINFO":
                        if (ws.Cells[cella.Item1, cella.Item2].Value == "OK")
                            valore1 = 1;
                        else
                            valore1 = 2;
                        break;
                    default:
                        if (cella != null)
                            valore1 = ws.Cells[cella.Item1, cella.Item2].Value;
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
                entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' IdParametroD = " + calcolo["IdParametroD"];

                if (entitaParametro.Count > 0)
                    valore1 = entitaParametro[0]["Valore"];
            }


        }

        #endregion
    }
}
