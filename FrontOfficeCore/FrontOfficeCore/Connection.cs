using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.IO;

namespace Iren.FrontOffice.Core
{
    //singleton
    public class Connection
    {
        //#region Variabili

        //private static Connection _conn;
        //private static string _connStr;
        //private static SqlConnection _sqlConn;
        //private static ConnectionState _state = ConnectionState.Closed;

        //#endregion

        //#region Proprietà

        //public static Connection Instance 
        //{
        //    get
        //    {
        //        if (_conn == null)
        //        {
        //            _conn = new Connection();
        //            //_conn.OpenConnection();
        //        }
        //        return _conn;
        //    }
        //}

        //#endregion

        //#region Costruttori

        //private Connection() {
        //    //la prima volta che viene lanciata la dll controlla che i parametri siano protetti
        //    //e nell'eventualità li protegge
        //    //CryptSection(System.Reflection.Assembly.GetExecutingAssembly());
        //}

        //#endregion

        //#region Metodi Statici

        //public static string GetConnStr() 
        //{
        //    return _connStr;
        //}

        //public static void SetConnStr(string name) 
        //{
        //    try
        //    {
        //        _connStr = ConfigurationManager.ConnectionStrings[name].ConnectionString;
        //    }
        //    catch { }
        //}

        //#endregion

        //#region Metodi Pubblici

        ////cripta i dati di connessione se sono in chiaro
        //public static void CryptSection(string location) 
        //{
        //    ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
        //    fileMap.ExeConfigFilename = location;
        //    var config = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);

        //    string provider = "RsaProtectedConfigurationProvider";
        //    ConfigurationSection connStrings = config.ConnectionStrings;
        //    if (connStrings != null)
        //    {
        //        if (!connStrings.SectionInformation.IsProtected)
        //        {
        //            if (!connStrings.ElementInformation.IsLocked)
        //            {
        //                connStrings.SectionInformation.ProtectSection(provider);

        //                connStrings.SectionInformation.ForceSave = true;
        //                config.Save(ConfigurationSaveMode.Modified);
        //            }
        //        }
        //    }
        //}

        //public static ConnectionState GetConnectionState() 
        //{
        //    return _state;
        //}

        //public static SqlConnection OpenConnection() 
        //{
        //    if (_connStr != null)
        //    {
        //        return OpenConnection(_connStr); ;
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

        //public static SqlConnection OpenConnection(string connectionString) 
        //{
        //    if (_sqlConn == null || _sqlConn.State == ConnectionState.Closed)
        //    {
        //        try
        //        {
        //            if (_sqlConn == null)
        //            {
        //                _connStr = connectionString;
        //                _sqlConn = new SqlConnection(_connStr);
        //                _sqlConn.StateChange += ConnectionStateChange;
        //            }

        //            _sqlConn.Open();
        //        }
        //        catch (Exception)
        //        {
        //            _connStr = null;
        //            return null;
        //        }
        //    }
        //    return _sqlConn;
        //}

        //public static void CloseConnection() 
        //{
        //    if (_sqlConn != null && _sqlConn.State == ConnectionState.Open)
        //    {
        //        _sqlConn.Close();
        //        _sqlConn = null;
        //    }
        //}

        //#endregion

        //#region Metodi Privati

        //private static void ConnectionStateChange(object sender, StateChangeEventArgs e) 
        //{
        //    _state = _sqlConn.State;
        //}

        //#endregion
    }
}
