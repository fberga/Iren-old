﻿using Iren.ToolsExcel.Core;
using System;
using System.Collections.Generic;
using System.Configuration;

namespace Iren.ToolsExcel.Base
{
    public class Simboli
    {
        public const string UNION = ".";

        //public static string nomeFile = "";

        public static string nomeApplicazione = "";
        private static bool emergenzaForzata = false;
        public static bool EmergenzaForzata
        {
            get
            {
                return emergenzaForzata;
            }
            set
            {
                if (emergenzaForzata != value)
                {
                    emergenzaForzata = value;

                    bool screenUpdating = Utility.Workbook.Application.ScreenUpdating;
                    if (screenUpdating)
                        Utility.Workbook.Application.ScreenUpdating = false;

                    bool isProtected = Utility.Workbook.Main.ProtectContents;
                    if (isProtected)
                        Utility.Workbook.Main.Unprotect(pwd);

                    Riepilogo main = new Riepilogo(Utility.Workbook.Main);
                    if (value)
                        main.RiepilogoInEmergenza();
                    else
                        if (Utility.DataBase.OpenConnection())
                            main.UpdateData();

                    Utility.Workbook.AggiornaLabelStatoDB();

                    if (isProtected)
                        Utility.Workbook.Main.Protect(pwd);

                    if (screenUpdating)
                        Utility.Workbook.Application.ScreenUpdating = true;
                }
            }
        }

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
                Handler.ChangeModificaDati(modificaDati);
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
                Handler.ChangeAmbiente(ambiente);
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
                Handler.ChangeStatoDB(DataBase.NomiDB.SQLSERVER, sqlServerOnline);
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
                Handler.ChangeStatoDB(DataBase.NomiDB.IMP, impiantiOnline);
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
                Handler.ChangeStatoDB(DataBase.NomiDB.ELSAG, elsagOnline);
            }
        }

        private static string appID = "";
        public static string AppID
        {
            get { return appID; }
            set 
            {
                appID = value;
                Utility.DataBase.ChangeAppSettings("AppID", value);
                mercato = GetMercatoByAppID(value);
                Handler.ChangeMercatoAttivo(mercato);
            }
        }

        private static string mercato;
        public static string Mercato
        {
            get { return mercato; }
        }

        public static string GetMercatoByAppID(string id)
        {
            List<string> mercati = new List<string>(ConfigurationManager.AppSettings["Mercati"].Split('|'));
            List<string> appIDs = new List<string>(ConfigurationManager.AppSettings["AppIDMSD"].Split('|'));

            return mercati[appIDs.IndexOf(id)];
        }
        public static string GetAppIDByMercato(string mercato)
        {
            List<string> mercati = new List<string>(ConfigurationManager.AppSettings["Mercati"].Split('|'));
            List<string> appIDs = new List<string>(ConfigurationManager.AppSettings["AppIDMSD"].Split('|'));

            return appIDs[mercati.IndexOf(mercato)];
        }
        public static string GetMercatoPrec()
        {
            return GetMercatoPrec(mercato);
        }
        public static string GetMercatoPrec(string mercato)
        {
            List<string> mercati = new List<string>(ConfigurationManager.AppSettings["Mercati"].Split('|'));
            int index = mercati.IndexOf(mercato);
            if(index > 0)
                return mercati[index - 1];

            return null;
        }

        public static int[] rgbSfondo = { 183, 222, 232 };
        public static int[] rgbLinee = { 33, 89, 104 };
        public static int[] rgbTitolo = { 49, 133, 156 };

    }
}
