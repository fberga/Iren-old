﻿using Iren.PSO.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Linq;

namespace Iren.PSO.Base
{
    public static class Simboli
    {
        public static string LocalBasePath 
        { get { return @"%APPDATA%\PSO\"; } }
        public static string RemoteBasePath
        { get { return @"\\pc1009235\Applicazioni\PSO"; } }

        private readonly static Dictionary<int, string> _fileApplicazione = new Dictionary<int, string>()
        {
            {1, "OfferteMGP"},
            {2, "InvioProgrammi"},
            {3, "InvioProgrammi"},
            {4, "InvioProgrammi"},
            {5, "ProgrammazioneImpianti"},
            {6, "UnitCommitment"},
            {7, "PrezziMSD"},
            {8, "SistemaComandi"},
            {9, "OfferteMSD"},
            {10, "OfferteMB"},
            {11, "ValidazioneTL"},
            {12, "PrevisioneCT"},
            {13, "InvioProgrammi"},
            {14, "ValidazioneGAS"},
            {15, "PrevisioneGAS"}
        };

        public static Dictionary<int, string> FileApplicazione 
        { get { return _fileApplicazione; } }
        
        public const string DEV = "Dev";
        public const string TEST = "Test";
        public const string PROD = "Prod";

        public const string UNION = ".";

        public static string NomeApplicazione 
        { get; set; }
        
        private static bool _emergenzaForzata = false;
        public static bool EmergenzaForzata 
        {
            get
            {
                return _emergenzaForzata;
            }
            set
            {
                if (_emergenzaForzata != value)
                {
                    _emergenzaForzata = value;

                    bool screenUpdating = Workbook.ScreenUpdating;
                    if (screenUpdating)
                        Workbook.ScreenUpdating = false;

                    bool isProtected = Workbook.Main.ProtectContents;
                    if (isProtected)
                        Workbook.Main.Unprotect(Workbook.Password);

                    Riepilogo main = new Riepilogo(Workbook.Main);
                    if (value)
                        main.RiepilogoInEmergenza();
                    else
                        if (DataBase.OpenConnection())
                        {
                            main.UpdateData();
                            DataBase.CloseConnection();
                        }

                    Workbook.AggiornaLabelStatoDB();

                    if (isProtected)
                        Workbook.Main.Protect(Workbook.Password);

                    if (screenUpdating)
                        Workbook.ScreenUpdating = true;
                }
            }
        }

        private static bool _modificaDati = false;
        public static bool ModificaDati 
        { 
            get 
            { 
                return _modificaDati; 
            } 
            
            set 
            {
                _modificaDati = value;
                Handler.ChangeModificaDati(_modificaDati);
            }
        }

        private static bool _sqlServerOnline = true;
        public static bool SQLServerOnline
        {
            get
            {
                return _sqlServerOnline;
            }

            set
            {
                _sqlServerOnline = value;
                Handler.ChangeStatoDB(Core.DataBase.NomiDB.SQLSERVER, _sqlServerOnline);
            }
        }

        private static bool _impiantiOnline = true;
        public static bool ImpiantiOnline
        {
            get
            {
                return _impiantiOnline;
            }

            set
            {
                _impiantiOnline = value;
                Handler.ChangeStatoDB(Core.DataBase.NomiDB.IMP, _impiantiOnline);
            }
        }

        private static bool _elsagOnline = true;
        public static bool ElsagOnline
        {
            get
            {
                return _elsagOnline;
            }

            set
            {
                _elsagOnline = value;
                Handler.ChangeStatoDB(Core.DataBase.NomiDB.ELSAG, _elsagOnline);
            }
        }
        
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
        { get { return oreMSD; } }

        public static string GetMercatoPrec(string mercato)
        {
            var index = Workbook.Repository[DataBase.TAB.MERCATI].AsEnumerable()
                .Where(r => r["DesMercato"].Equals(mercato))
                .Select(r => Workbook.Repository[DataBase.TAB.MERCATI].Rows.IndexOf(r))
                .FirstOrDefault();

            if (index > 0)
                return Workbook.Repository[DataBase.TAB.MERCATI].Rows[index - 1]["DesMercato"].ToString();

            return "";

        }
    }
}
