using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Iren.PSO.Base
{
    public static class Simboli
    {
        public static string LocalBasePath 
        { get { return @"%APPDATA%\PSO\"; } }
        public static string RemoteBasePath
        {   /* Modifica per rliascio in Test ***** BEGIN ***** */
            get { return @"\\srvpso\Applicazioni\PSO_TEST"; } //TODO Riportare a PSO per rilasci in prod
          //   get { return @"\\srvpso\Applicazioni\PSO"; }
            /* Modifica per rliascio in Test ***** END ***** */
        } 

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
            {15, "PrevisioneGAS"},
            //TODO      
            // Modifica per InvioProgrammi MSD5 e MSD6 ***** BEGIN *****
            {16, "InvioProgrammi"},
            {17, "InvioProgrammi"},
            // Modifica per InvioProgrammi MSD5 e MSD6 ***** END *****
            {18, "OfferteMI"}
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

                    bool autoCalc = Workbook.Application.Calculation == Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic;

                    if (autoCalc)
                        Workbook.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;

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

                    if(autoCalc)
                        Workbook.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic;

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
            //TODO modificare in base ai nuovi orari
             // Modifica per InvioProgrammi MSD5 e MSD6 e nuovi orari ***** BEGIN *****
            {0, "MSD2"},
            {1, "MSD2"},
            {2, "MSD2"},
            {3, "MSD2"},
            {4, "MSD3"},
            {5, "MSD3"},
            {6, "MSD3"},
            {7, "MSD3"},
            {8, "MSD4"},
            {9, "MSD4"},
            {10, "MSD4"},
            {11, "MSD4"},
            {12, "MSD5"},
            {13, "MSD5"},
            {14, "MSD5"},
            {15, "MSD5"},
            {16, "MSD6"},
            {17, "MSD6"},
            {18, "MSD6"},
            {19, "MSD6"},
            {20, "MSD1"},
            {21, "MSD1"},
            {22, "MSD1"},
            {23, "MSD1"},
            // Modifica per InvioProgrammi MSD5 e MSD6 e nuovi orari ***** END *****

             /*
             // Orari fino al 01/02/2017
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
              */
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

        private readonly static Dictionary<string, MB> mercatiMB = new Dictionary<string, MB>()
        {
            //TODO modificare in base a nuovi orari
            {"MB1", new MB(0,1,8)},
            {"MB2", new MB(7,9,12)},
            {"MB3", new MB(11,13,16)},
            {"MB4", new MB(15,17,22)},
            {"MB5", new MB(21,23,25)}
        };

        public static Dictionary<string, MB> MercatiMB { get { return mercatiMB; } }

        public static int GetMarketOffset(int hour)
        {
            if (Workbook.Repository.Applicazione["ModificaDinamica"].Equals("1"))
            {
                int offset = Simboli.MercatiMB["MB1"].Fine;
                if (hour >= Simboli.MercatiMB["MB2"].Chiusura)
                {
                    string mercatoChiuso = Simboli.MercatiMB
                        .Where(kv => kv.Value.Chiusura <= hour)
                        .Select(kv => kv.Key)
                        .Last();
                    offset = Simboli.MercatiMB[mercatoChiuso].Fine;
                }

                return offset;
            }
            return 0;
        }

        public static Range GetMarketCompleteRange(string mercato, DateTime giorno, Range rng)
        {
            if (!mercatiMB.ContainsKey(mercato))
                return null;

            int[] orario = new int[2] { Simboli.MercatiMB[mercato].Inizio, Math.Min(Simboli.MercatiMB[mercato].Fine, Date.GetOreGiorno(giorno)) };

            return new Range(rng.StartRow, rng.StartColumn + orario[0] - 1, 1, orario[1] - orario[0] + 1);
        }
    }
}

public class MB
{
    public int Chiusura { get; private set; }
    public int Inizio { get; private set; }
    public int Fine { get; private set; }

    public MB(int chiusura, int inizio, int fine)
    {
        Chiusura = chiusura;
        Inizio = inizio;
        Fine = fine;
    }
}
