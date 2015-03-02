﻿
using Iren.ToolsExcel.Core;
namespace Iren.ToolsExcel.Base
{
    public class Simboli
    {
        public const string UNION = ".", 
            ALL = "ALL";

        public static string nomeApplicazione = "";
        public static int intervalloGiorni = 0;

        public static string pwd = "";

        private static bool modificaDati = false;
        public static bool ModificaDati 
        { 
            get 
            { 
                return modificaDati; 
            } 
            
            set 
            {
                modificaDati = value;
                BaseHandler.ChangeModificaDati(modificaDati);
            }
        }

        private static string ambiente = "";
        public static string Ambiente
        {
            get
            {
                return ambiente;
            }

            set
            {
                ambiente = value;
                BaseHandler.ChangeAmbiente(ambiente);
            }
        }

        private static bool sqlServerOnline = true;
        public static bool SQLServerOnline
        {
            get
            {
                return sqlServerOnline;
            }

            set
            {
                sqlServerOnline = value;
                BaseHandler.ChangeStatoDB(DataBase.NomiDB.SQLSERVER, sqlServerOnline);
            }
        }

        private static bool impiantiOnline = true;
        public static bool ImpiantiOnline
        {
            get
            {
                return impiantiOnline;
            }

            set
            {
                impiantiOnline = value;
                BaseHandler.ChangeStatoDB(DataBase.NomiDB.IMP, impiantiOnline);
            }
        }

        private static bool elsagOnline = true;
        public static bool ElsagOnline
        {
            get
            {
                return elsagOnline;
            }

            set
            {
                elsagOnline = value;
                BaseHandler.ChangeStatoDB(DataBase.NomiDB.ELSAG, elsagOnline);
            }
        }

        public const string NameSpace = "Iren.ToolsExcel";

    }
}