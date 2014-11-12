using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace RiMoST2
{
    class Command
    {
        #region Costruttori

        public Command() {}

        #endregion

        #region Metodi

        public SqlCommand SqlCmd(string commandText, CommandType commandType, int timeout = 300)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = Connection.Instance.OpenConnection();
            cmd.CommandText = commandText;
            cmd.CommandType = commandType;
            cmd.CommandTimeout = timeout;
            return cmd;
        }
        public SqlCommand SqlCmd(string commandText)
        {
            return SqlCmd(commandText, CommandType.StoredProcedure);
        }
        
        public SqlCommand SqlCmd(string commandText, CommandType commandType, Dictionary<string, object> parameters, int timeout = 300)
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
            catch (Exception ex)
            {                
            }
            return cmd;
        }
        public SqlCommand SqlCmd(string commandText, Dictionary<String, Object> parameters)
        {
            return SqlCmd(commandText, CommandType.StoredProcedure, parameters);
        }
       
        #endregion

    }
}
