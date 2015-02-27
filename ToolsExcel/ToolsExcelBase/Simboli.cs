
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

        public const string NameSpace = "Iren.ToolsExcel";

    }
}
