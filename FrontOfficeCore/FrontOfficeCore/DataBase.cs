using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

using DataTable = System.Data.DataTable;
using System.Data;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Globalization;

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

        public static string DataAttiva { get { return _dataAttiva; } set { _dataAttiva = value; } }
        public static DateTime Data { get { return DateTime.ParseExact(_dataAttiva, "yyyyMMdd", CultureInfo.InvariantCulture); } }
        public static int IdUtenteAttivo { get { return _idUtenteAttivo; } }
        public static int IdApplicazione { get { return _idApplicazione; } }

        #endregion

        #region Costruttori

        public DataBase(string dbName)
        {
            _cmd = new Command();
            Connection.CloseConnection();
            Connection.SetConnStr(dbName);            
        }

        //public DataBase() {}

        #endregion

        #region Metodi

        public Dictionary<NomiDB, ConnectionState> StatoDB()
        {
            _statoDB[NomiDB.SQLSERVER] = Connection.Instance.GetConnectionState();

            if (_statoDB[NomiDB.SQLSERVER] == ConnectionState.Open)
            {
                DataView imp = Select("spCheckDB", "@Nome=IMP").DefaultView;
                DataView elsag = Select("spCheckDB", "@Nome=ELSAG").DefaultView;

                if (imp.Count > 0 && imp[0]["Stato"].Equals(0))
                {
                    _statoDB[NomiDB.IMP] = ConnectionState.Open;
                }
                if (elsag.Count > 0 && elsag[0]["Stato"].Equals(0))
                {
                    _statoDB[NomiDB.ELSAG] = ConnectionState.Open;
                }
            }
            return _statoDB;
        }

        public void setParameters(string dataAttiva, int idUtenteAttivo, int idApplicazione)
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

        public DataTable Select(string storedProcedure, String parameters)
        {
            return Select(storedProcedure, getParamsFromString(parameters));
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

        public DataTable Select(string storedProcedure)
        {
            QryParams parameters = new QryParams();
            return Select(storedProcedure, parameters);
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

        public System.Version GetCurrentV()
        {
            return Assembly.GetExecutingAssembly().GetName().Version;
        }

        #endregion
    }
}
