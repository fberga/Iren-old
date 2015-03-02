using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Iren.ToolsExcel.Core
{
    public class DataBase : INotifyPropertyChanged
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

        public struct SP
        {
            public const string APPLICAZIONE = "spApplicazioneProprieta",
                APPLICAZIONE_INFORMAZIONE = "spApplicazioneInformazione",
                APPLICAZIONE_INFORMAZIONE_COMMENTO = "spApplicazioneInformazioneCommento",
                APPLICAZIONE_INIT = "spApplicazioneInit",
                APPLICAZIONE_LOG = "spApplicazioneLog",
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
                ENTITA_AZIONE_INFORMAZIONE = "spEntitaAzioneInformazione",
                ENTITACALCOLO = "spEntitaCalcolo",
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
                INSERT_LOG = "spInsertLog",
                INSERT_PROGRAMMAZIONE_PARAMETRO = "spInsertProgrammazione_Parametro",
                TIPOLOGIA_CHECK = "spTipologiaCheck",
                TIPOLOGIA_RAMPA = "spTipologiaRampa",
                UTENTE = "spUtente";
        }

        #endregion

        #region Variabili

        private Command _cmd;
        private Command _internalCmd;

        private System.Threading.Timer checkDBTrhead;

        private SqlConnection _sqlConn;
        private SqlConnection _internalsqlConn;
        private string _connStr = "";
        private ConnectionState _state = ConnectionState.Closed;

        private bool _rightClosure;

        private string _dataAttiva = "";
        private int _idUtenteAttivo = -1;
        private int _idApplicazione = -1;
        private Dictionary<NomiDB, ConnectionState> _statoDB = new Dictionary<NomiDB, ConnectionState>() { 
            {NomiDB.SQLSERVER, ConnectionState.Closed},
            {NomiDB.IMP, ConnectionState.Closed},
            {NomiDB.ELSAG, ConnectionState.Closed}
        };

        #endregion

        #region Proprietà

        public DateTime DataAttiva { get { return DateTime.ParseExact(_dataAttiva, "yyyyMMdd", CultureInfo.InvariantCulture); } }
        public int IdUtenteAttivo { get { return _idUtenteAttivo; } }
        public int IdApplicazione { get { return _idApplicazione; } }

        public Dictionary<NomiDB, ConnectionState> StatoDB { get { return _statoDB; } }

        #endregion

        #region Costruttori

        public DataBase(string dbName)
        {
            try
            {
                _connStr = ConfigurationManager.ConnectionStrings[dbName].ConnectionString;
                _sqlConn = new SqlConnection(_connStr);
                _internalsqlConn = new SqlConnection(_connStr);

                checkDBTrhead = new System.Threading.Timer(CheckDB, null, 0, 1000 * 60);

                //_sqlConn.StateChange += ConnectionStateChange;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Core.DataBase - ERROR!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            _cmd = new Command(_sqlConn);
            _internalCmd = new Command(_internalsqlConn);
        }

        #endregion

        #region Metodi Pubblici

        //public Dictionary<NomiDB, ConnectionState> StatoDB()
        //{
        //    OpenConnection();

        //    _statoDB[NomiDB.SQLSERVER] = _sqlConn.State;

        //    if (_statoDB[NomiDB.SQLSERVER] == ConnectionState.Open)
        //    {
        //        DataView imp = Select("spCheckDB", "@Nome=IMP", 3).DefaultView;
        //        //se va in timeout la connessione si chiude
        //        OpenConnection();
        //        DataView elsag = Select("spCheckDB", "@Nome=ELSAG", 3).DefaultView;
        //        //se va in timeout la connessione si chiude
        //        OpenConnection();

        //        if (imp.Count > 0 && imp[0]["Stato"].Equals(0))
        //        {
        //            _statoDB[NomiDB.IMP] = ConnectionState.Open;
        //        }
        //        else
        //        {
        //            _statoDB[NomiDB.IMP] = ConnectionState.Closed;
        //        }

        //        if (elsag.Count > 0 && elsag[0]["Stato"].Equals(0))
        //        {
        //            _statoDB[NomiDB.ELSAG] = ConnectionState.Open;
        //        }
        //        else
        //        {
        //            _statoDB[NomiDB.ELSAG] = ConnectionState.Closed;
        //        }
        //    }

        //    return _statoDB;
        //}

        public bool OpenConnection()
        {
            return OpenConnection(_sqlConn);
        }
        public bool CloseConnection()
        {
            return CloseConnection(_sqlConn);
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

            try
            {
                _cmd.SqlCmd(storedProcedure, parameters).ExecuteNonQuery();
            }
            catch (TimeoutException) { }
            
        }
        public void InsertLog(TipologiaLOG tipologia, string messaggio)
        {
            QryParams logParam = new QryParams()
            {
                {"@IdTipologia", tipologia},
                {"@Messaggio", messaggio}
            };

            Insert(SP.INSERT_LOG, logParam);
        }

        public DataTable Select(string storedProcedure, QryParams parameters, int timeout = 300)
        {
            return Select(_cmd, storedProcedure, parameters, timeout);
        }
        public DataTable Select(string storedProcedure, String parameters, int timeout = 300)
        {
            return Select(storedProcedure, getParamsFromString(parameters), timeout);
        }
        public DataTable Select(string storedProcedure, int timeout = 300)
        {
            QryParams parameters = new QryParams();
            return Select(storedProcedure, parameters, timeout);
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
        }

        #endregion

        #region Metodi Privati

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

        private bool OpenConnection(SqlConnection conn)
        {
            try
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }
        private bool CloseConnection(SqlConnection conn)
        {
            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        private DataTable Select(Command cmd, string storedProcedure, QryParams parameters, int timeout = 300)
        {
            if (!parameters.ContainsKey("@IdApplicazione") && _idApplicazione != -1)
                parameters.Add("@IdApplicazione", _idApplicazione);
            if (!parameters.ContainsKey("@IdUtente") && _idUtenteAttivo != -1)
                parameters.Add("@IdUtente", _idUtenteAttivo);
            if (!parameters.ContainsKey("@Data") && _dataAttiva != "")
                parameters.Add("@Data", _dataAttiva);
            try
            {
                using (SqlDataReader dr = cmd.SqlCmd(storedProcedure, parameters, timeout).ExecuteReader())
                {
                    DataTable dt = new DataTable();
                    dt.Load(dr);

                    return dt;
                }
            }
            catch (SqlException)
            {
                return new DataTable();
            }

        }
        private DataTable Select(Command cmd, string storedProcedure, String parameters, int timeout = 300)
        {
            return Select(cmd, storedProcedure, getParamsFromString(parameters), timeout);
        }

        #endregion

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged(String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }


        private void CheckDB(object state)
        {
            Dictionary<NomiDB, ConnectionState> oldStatoDB = new Dictionary<NomiDB,ConnectionState>(_statoDB);
            if (OpenConnection(_internalsqlConn))
            {
                _statoDB[NomiDB.SQLSERVER] = _internalsqlConn.State;

                if (_statoDB[NomiDB.SQLSERVER] == ConnectionState.Open)
                {
                    DataView imp = Select(_internalCmd, "spCheckDB", "@Nome=IMP", 3).DefaultView;
                    //se va in timeout la connessione si chiude
                    OpenConnection(_internalsqlConn);
                    DataView elsag = Select(_internalCmd, "spCheckDB", "@Nome=ELSAG", 3).DefaultView;
                    //se va in timeout la connessione si chiude
                    OpenConnection(_internalsqlConn);

                    if (imp.Count > 0 && imp[0]["Stato"].Equals(0))
                        _statoDB[NomiDB.IMP] = ConnectionState.Open;
                    else
                        _statoDB[NomiDB.IMP] = ConnectionState.Closed;

                    if (elsag.Count > 0 && elsag[0]["Stato"].Equals(0))
                        _statoDB[NomiDB.ELSAG] = ConnectionState.Open;
                    else
                        _statoDB[NomiDB.ELSAG] = ConnectionState.Closed;
                }

                if (_statoDB[NomiDB.SQLSERVER] != oldStatoDB[NomiDB.SQLSERVER]
                    || _statoDB[NomiDB.IMP] != oldStatoDB[NomiDB.IMP]
                    || _statoDB[NomiDB.ELSAG] != oldStatoDB[NomiDB.ELSAG])
                {
                    NotifyPropertyChanged("StatoDB");
                }
            }
        }
    }
}
