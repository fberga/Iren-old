using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Iren.FrontOffice.Base
{
    public class Simboli
    {
        public const string UNION = ".", 
            ALL = "ALL";

        public static string nomeApplicazione = "";
        public static int intervalloGiorni = 0;

        private static bool modificaDati = false;
        public static string pwd = "";
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

    }
}
