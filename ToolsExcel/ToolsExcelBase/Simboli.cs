﻿using Iren.ToolsExcel.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Linq;

namespace Iren.ToolsExcel.Base
{
    public class Simboli
    {
        public const string DEV = "Dev";
        public const string TEST = "Test";
        public const string PROD = "Prod";

        public const string UNION = ".";

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

                    bool screenUpdating = Utility.Workbook.ScreenUpdating;
                    if (screenUpdating)
                        Utility.Workbook.ScreenUpdating = false;

                    bool isProtected = Utility.Workbook.Main.ProtectContents;
                    if (isProtected)
                        Utility.Workbook.Main.Unprotect(Utility.Workbook.Password);

                    Riepilogo main = new Riepilogo(Utility.Workbook.Main);
                    if (value)
                        main.RiepilogoInEmergenza();
                    else
                        if (Utility.DataBase.OpenConnection())
                        {
                            main.UpdateData();
                            Utility.DataBase.CloseConnection();
                        }

                    Utility.Workbook.AggiornaLabelStatoDB();

                    if (isProtected)
                        Utility.Workbook.Main.Protect(Utility.Workbook.Password);

                    if (screenUpdating)
                        Utility.Workbook.ScreenUpdating = true;
                }
            }
        }

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

        //private static string mercato;
        //public static string Mercato
        //{
        //    get { return mercato; }
        //}

        //public static string GetMercatoByAppID(string id)
        //{
        //    List<string> mercati = new List<string>(Utility.Workbook.AppSettings("Mercati").Split('|'));
        //    List<string> appIDs = new List<string>(Utility.Workbook.AppSettings("AppIDMSD").Split('|'));

        //    return mercati[appIDs.IndexOf(id)];
        //}
        //public static int GetAppIDByMercato(string mercato)
        //{
        //    List<string> mercati = new List<string>(Utility.Workbook.AppSettings("Mercati").Split('|'));
        //    List<string> appIDs = new List<string>(Utility.Workbook.AppSettings("AppIDMSD").Split('|'));

        //    List<int> ids = new List<int>();

        //    foreach (string id in appIDs)
        //        ids.Add(int.Parse(id));

        //    return ids[mercati.IndexOf(mercato)];
        //}
        //public static string GetMercatoPrec()
        //{
        //    return GetMercatoPrec(mercato);
        //}
        //public static string GetMercatoPrec(string mercato)
        //{
        //    List<string> mercati = new List<string>(Utility.Workbook.AppSettings("Mercati").Split('|'));
        //    int index = mercati.IndexOf(mercato);
        //    if(index > 0)
        //        return mercati[index - 1];

        //    return null;
        //}

        //public static string Stagione
        //{
        //    get { return GetStagione(); }
        //    set
        //    {
        //        Utility.Workbook.Ids
                
        //        string idStagione = GetIdStagione(value);
        //        Utility.Workbook.ChangeAppSettings("Stagione", idStagione);
        //        DefinedNames definedNames = new DefinedNames("Previsione");
        //        DateTime dataFine = Utility.Workbook.DataAttiva.AddDays(Struct.intervalloGiorni);
        //        Range rng = definedNames.Get("CT_TORINO", "STAGIONE", Utility.Date.SuffissoDATA1, Utility.Date.GetSuffissoOra(1)).Extend(colOffset: Utility.Date.GetOreIntervallo(dataFine));
        //        Utility.Workbook.Sheets["Previsione"].Range[rng.ToString()].Value = idStagione;
        //    }
        //}

        //private static string GetIdStagione(string stagione) 
        //{
        //    List<string> stagioni = new List<string>(Utility.Workbook.AppSettings("Stagioni").Split('|'));
        //    List<string> idStagioni = new List<string>(Utility.Workbook.AppSettings("IdStagioni").Split('|'));

        //    return idStagioni[stagioni.IndexOf(stagione)];
        //}
        //public static string GetStagione(string id)
        //{
        //    List<string> stagioni = new List<string>(Utility.Workbook.AppSettings("Stagioni").Split('|'));
        //    List<string> idStagioni = new List<string>(Utility.Workbook.AppSettings("IdStagioni").Split('|'));

        //    return stagioni[idStagioni.IndexOf(id)];
        //}
        //private static string GetStagione()
        //{
        //    return GetStagione(Utility.Workbook.AppSettings("Stagione"));
        //}
        
        public static int[] rgbSfondo = { 228, 144, 144 };
        public static int[] rgbLinee = { 176, 0, 0 };
        public static int[] rgbTitolo = { 206, 58, 58 };

        private readonly static Dictionary<int, string> oreMSD = new Dictionary<int, string>() 
        { 
            {0, "MSD1"},
            {1, "MSD1"},
            {2, "MSD1"},
            {3, "MSD1"},
            {4, "MSD2"},
            {5, "MSD2"},
            {6, "MSD2"},
            {7, "MSD2"},
            {8, "MSD3"},
            {9, "MSD3"},
            {10, "MSD3"},
            {11, "MSD3"},
            {12, "MSD4"},
            {13, "MSD4"},
            {14, "MSD4"},
            {15, "MSD4"},
            {16, "MSD4"},
            {17, "MSD4"},
            {18, "MSD4"},
            {19, "MSD1"},
            {20, "MSD1"},
            {21, "MSD1"},
            {22, "MSD1"},
            {23, "MSD1"},
        };


        public static Dictionary<int, string> OreMSD
        {
            get
            {
                return oreMSD;
            }
        }

        public static string GetMercatoPrec(string mercato)
        {
            var index = Utility.Workbook.Repository[Utility.DataBase.TAB.MERCATI].AsEnumerable()
                .Where(r => r["DesMercato"].Equals(mercato))
                .Select(r => Utility.Workbook.Repository[Utility.DataBase.TAB.MERCATI].Rows.IndexOf(r))
                .FirstOrDefault();

            if (index > 0)
                return Utility.Workbook.Repository[Utility.DataBase.TAB.MERCATI].Rows[index - 1]["DesMercato"].ToString();

            return "";

        }
    }
}
