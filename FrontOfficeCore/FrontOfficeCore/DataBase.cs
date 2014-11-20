using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

using DataTable = System.Data.DataTable;
using System.Data;
using System.Reflection;

namespace Iren.FrontOffice.Core
{
    public class DataBase
    {
        #region Variabili

        private Command _cmd;

        private static string _dataAttiva = "";
        private static int _idUtenteAttivo = -1;
        private static int _idApplicazione = -1;

        #endregion

        #region Proprietà

        public static string DataAttiva { get { return _dataAttiva; } set { _dataAttiva = value; } }
        public static int IdUtenteAttivo { get { return _idUtenteAttivo; } }
        public static int IdApplicazione { get { return _idApplicazione; } }

        #endregion

        #region Costruttori

        public DataBase(string dbName)
        {
            _cmd = new Command();
            Connection.SetConnStr(dbName);
        }

        #endregion

        #region Metodi

        public void setParameters(string dataAttiva, int idUtenteAttivo, int idApplicazione)
        {
            _dataAttiva = dataAttiva;
            _idUtenteAttivo = idUtenteAttivo;
            _idApplicazione = idApplicazione;
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

        public System.Version GetCurrentV()
        {
            return Assembly.GetExecutingAssembly().GetName().Version;
        }

        #endregion
    }
}
