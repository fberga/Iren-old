﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Iren.FrontOffice.Core
{
    public class DataBase
    {
        #region Nomi di Sistema

        public enum TipologiaLOG
        {
            LogErrore = 1,
            LogCarica = 2,
            LogGenera = 3,
            LogEsporta = 4,
            LogModifica = 5,
            LogAccesso = 6
        }

        public enum NomiDB
        {
            SQLSERVER = 1,
            IMP = 2,
            ELSAG = 3
        }

        public const string ALL = "ALL";

        public struct StoredProcedure
        {
            public const string UTENTE = "spUtente",
            APPLICAZIONE = "spApplicazioneProprieta",
            GETVERSION = "spGetVersione",
            LOG = "spLog",
            INSERT_LOG = "spInsertLog",
            APP_INFO = "spApplicazioneInformazione",
            AZIONE = "spAzione",
            CATEGORIA = "spCategoria",
            AZIONECATEGORIA = "spAzioneCategoria",
            ENTITAAZIONE = "spEntitaAzione",
            ENTITAINFORMAZIONE = "spEntitaInformazione",
            ENTITAAZIONEINFORMAZIONE = "spEntitaAzioneInformazione",
            CALCOLO = "spCalcolo",
            CALCOLOINFORMAZIONE = "spCalcoloInformazione",
            ENTITACALCOLO = "spEntitaCalcolo",
            ENTITAGRAFICO = "spEntitaGrafico",
            ENTITAGRAFICOINFORMAZIONE = "spEntitaGraficoInformazione",
            ENTITACOMMITMENT = "spEntitaCommitment",
            ENTITARAMPA = "spEntitaRampa",
            ENTITAASSETTO = "spEntitaAssetto",
            ENTITAPROPRIETA = "spEntitaProprieta",
            ENTITAINFORMAZIONEFORMATTAZIONE = "spEntitaInformazioneFormattazione",
            TIPOLOGIACHECK = "spTipologiaCheck",
            TIPOLOGIARAMPA = "spTipologiaRampa",
            CATEGORIAENTITA = "spCategoriaEntita",
            APP_RIEPILOGO = "spApplicazioneRiepilogo",
            INS_PROG_PARAM = "spInsertProgrammazione_Parametro",
            CHECK_MOD_STRUCT = "spCheckModificaStruttura",
            ENTITAPARAMETROD = "spEntitaParametroD",
            ENTITAPARAMETROH = "spEntitaParametroH";
        }

        #endregion

        #region Variabili

        private Command _cmd;

        private static SqlConnection _sqlConn;
        private static string _connStr = "";
        private static ConnectionState _state = ConnectionState.Closed;

        private bool _rightClosure;

        private static string _dataAttiva = "";
        private static int _idUtenteAttivo = -1;
        private static int _idApplicazione = -1;
        private static Dictionary<NomiDB, ConnectionState> _statoDB = new Dictionary<NomiDB, ConnectionState>() { 
            {NomiDB.SQLSERVER, ConnectionState.Closed},
            {NomiDB.IMP, ConnectionState.Closed},
            {NomiDB.ELSAG, ConnectionState.Closed}
        };

        #endregion

        #region Proprietà

        public static DateTime DataAttiva { get { return DateTime.ParseExact(_dataAttiva, "yyyyMMdd", CultureInfo.InvariantCulture); } }
        public static int IdUtenteAttivo { get { return _idUtenteAttivo; } }
        public static int IdApplicazione { get { return _idApplicazione; } }

        #endregion

        #region Costruttori

        public DataBase(string dbName)
        {
            try
            {
                _connStr = ConfigurationManager.ConnectionStrings[dbName].ConnectionString;
                _sqlConn = new SqlConnection(_connStr);
                _sqlConn.StateChange += ConnectionStateChange;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Core.DataBase - ERROR!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            _cmd = new Command(_sqlConn);                        
        }

        #endregion

        #region Metodi Pubblici

        public Dictionary<NomiDB, ConnectionState> StatoDB()
        {
            _statoDB[NomiDB.SQLSERVER] = _sqlConn.State;

            if (_statoDB[NomiDB.SQLSERVER] == ConnectionState.Open)
            {
                DataView imp = Select("spCheckDB", "@Nome=IMP").DefaultView;
                DataView elsag = Select("spCheckDB", "@Nome=ELSAG").DefaultView;

                if (imp.Count > 0 && imp[0]["Stato"].Equals(0))
                {
                    _statoDB[NomiDB.IMP] = ConnectionState.Open;
                }
                else
                {
                    _statoDB[NomiDB.IMP] = ConnectionState.Closed;
                }

                if (elsag.Count > 0 && elsag[0]["Stato"].Equals(0))
                {
                    _statoDB[NomiDB.ELSAG] = ConnectionState.Open;
                }
                else
                {
                    _statoDB[NomiDB.ELSAG] = ConnectionState.Closed;
                }
            }            

            return _statoDB;
        }

        public bool OpenConnection()
        {
            try
            {
                if(_sqlConn.State == ConnectionState.Closed)
                    _sqlConn.Open();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Core.DataBase - ERROR!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }

            return true;
        }
        public bool CloseConnection()
        {
            try
            {
                if (_sqlConn.State == ConnectionState.Open)
                {
                    _rightClosure = true;
                    _sqlConn.Close();
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Core.DataBase - ERROR!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        public void SetParameters(string dataAttiva, int idUtenteAttivo, int idApplicazione)
        {
            _dataAttiva = dataAttiva;
            _idUtenteAttivo = idUtenteAttivo;
            _idApplicazione = idApplicazione;
        }
        public void ChangeDate(string dataAttiva)
        {
            _dataAttiva = dataAttiva;
        }

        public void Insert(string storedProcedure, QryParams parameters)
        {
            if (!parameters.ContainsKey("@IdApplicazione") && _idApplicazione != -1)
                parameters.Add("@IdApplicazione", _idApplicazione);
            if (!parameters.ContainsKey("@IdUtente") && _idUtenteAttivo != -1)
                parameters.Add("@IdUtente", _idUtenteAttivo);
            if (!parameters.ContainsKey("@Data") && _dataAttiva != "")
                parameters.Add("@Data", _dataAttiva);

            _cmd.SqlCmd(storedProcedure, parameters).ExecuteNonQuery();
        }
        public void InsertLog(TipologiaLOG tipologia, string messaggio)
        {
            QryParams logParam = new QryParams()
            {
                {"@IdTipologia", tipologia},
                {"@Messaggio", messaggio}
            };

            Insert(StoredProcedure.INSERT_LOG, logParam);
        }

        public DataTable Select(string storedProcedure, QryParams parameters)
        {
            if (!parameters.ContainsKey("@IdApplicazione") && _idApplicazione != -1)
                parameters.Add("@IdApplicazione", _idApplicazione);
            if (!parameters.ContainsKey("@IdUtente") && _idUtenteAttivo != -1)
                parameters.Add("@IdUtente", _idUtenteAttivo);
            if (!parameters.ContainsKey("@Data") && _dataAttiva != "")
                parameters.Add("@Data", _dataAttiva);

            using (SqlDataReader dr = _cmd.SqlCmd(storedProcedure, parameters).ExecuteReader())
            {
                DataTable dt = new DataTable();
                dt.Load(dr);

                return dt;
            }
        }
        public DataTable Select(string storedProcedure, String parameters)
        {
            return Select(storedProcedure, getParamsFromString(parameters));
        }
        public DataTable Select(string storedProcedure)
        {
            QryParams parameters = new QryParams();
            return Select(storedProcedure, parameters);
        }

        public System.Version GetCurrentV()
        {
            return Assembly.GetExecutingAssembly().GetName().Version;
        }

        //cripta i dati di connessione se sono in chiaro
        //public static void CryptSection(string location)
        public static void CryptSection()
        {
            //ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
            //fileMap.ExeConfigFilename = location;
            //var config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            string provider = "RsaProtectedConfigurationProvider";
            ConfigurationSection connStrings = config.ConnectionStrings;
            if (connStrings != null)
            {
                if (!connStrings.SectionInformation.IsProtected)
                {
                    if (!connStrings.ElementInformation.IsLocked)
                    {
                        connStrings.SectionInformation.ProtectSection(provider);

                        connStrings.SectionInformation.ForceSave = true;
                        config.Save(ConfigurationSaveMode.Modified);
                    }
                }
            }

            ConfigurationSection appSettings = config.AppSettings;
            if (appSettings != null)
            {
                if(!appSettings.SectionInformation.IsProtected)
                {
                    if (!appSettings.ElementInformation.IsLocked)
                    {
                        appSettings.SectionInformation.ProtectSection(provider);

                        appSettings.SectionInformation.ForceSave = true;
                        config.Save(ConfigurationSaveMode.Modified);
                    }
                }
            }
        }

        #endregion

        #region Metodi Privati

        private void ConnectionStateChange(object sender, StateChangeEventArgs e)
        {            
            if (e.OriginalState == ConnectionState.Open && e.CurrentState == ConnectionState.Closed && !_rightClosure)
                System.Windows.Forms.MessageBox.Show("Attenzione, la connessione al DB si è chiusa in modo inaspettato...", "Core.DataBase - ERROR!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

            _state = _sqlConn.State;
            _rightClosure = false;
        }

        private QryParams getParamsFromString(string parameters)
        {
            Regex regex = new Regex(@"@\w+[=:][^;:=]+");
            MatchCollection paramsList = regex.Matches(parameters);
            Regex split = new Regex("[=:]");
            QryParams o = new QryParams();
            foreach (Match par in paramsList)
            {
                string[] keyVal = split.Split(par.Value);
                if (keyVal.Length != 2)
                    throw new InvalidExpressionException("The provided list of parameters is invalid.");
                o.Add(keyVal[0], keyVal[1]);
            }
            return o;
        }

        #endregion
    }
}
