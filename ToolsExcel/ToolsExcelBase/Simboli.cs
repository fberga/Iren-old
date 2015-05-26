using Iren.ToolsExcel.Core;

namespace Iren.ToolsExcel.Base
{
    public class Simboli
    {
        public const string UNION = ".";

        public static string nomeFile = "";

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
                emergenzaForzata = value;
                Utility.Workbook.AggiornaLabelStatoDB();
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

        //public const string NameSpace = "Iren.ToolsExcel";

        public static int[] rgbSfondo = { 183, 222, 232 };
        public static int[] rgbLinee = { 33, 89, 104 };
        public static int[] rgbTitolo = { 49, 133, 156 };

    }
}
