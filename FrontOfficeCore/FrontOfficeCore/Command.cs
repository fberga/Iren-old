using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace Iren.FrontOffice.Core
{
    class Command
    {
        #region Variabili

        private SqlConnection _sqlConn;

        #endregion

        #region Costruttori

        public Command(SqlConnection sqlConn) 
        {
            _sqlConn = sqlConn;
        }

        #endregion

        #region Metodi

        public SqlCommand SqlCmd(string commandText, CommandType commandType, int timeout = 300)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = _sqlConn;
            cmd.CommandText = commandText;
            cmd.CommandType = commandType;
            cmd.CommandTimeout = timeout;
            return cmd;
        }
        public SqlCommand SqlCmd(string commandText)
        {
            return SqlCmd(commandText, CommandType.StoredProcedure);
        }
        public SqlCommand SqlCmd(string commandText, CommandType commandType, QryParams parameters, int timeout = 300)
        {
            SqlCommand cmd = SqlCmd(commandText, commandType, timeout);
            try
            {
                SqlCommandBuilder.DeriveParameters(cmd);
                foreach (SqlParameter par in cmd.Parameters)
                {
                    if(parameters.ContainsKey(par.ParameterName))
                        par.Value = parameters[par.ParameterName];
                }
            }
            catch (Exception)
            {                
            }
            return cmd;
        }
        public SqlCommand SqlCmd(string commandText, QryParams parameters)
        {
            return SqlCmd(commandText, CommandType.StoredProcedure, parameters);
        }
       
        #endregion

    }
}
