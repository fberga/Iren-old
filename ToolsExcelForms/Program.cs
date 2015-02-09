using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace FrontOfficeCOMTools
{
    static class Program
    {
        /// <summary>
        /// Punto di ingresso principale dell'applicazione.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Match func = Regex.Match("ROUND1", @"abs|floor|round\d*|power|max(_H)?|min(_H)?|sum|avg", RegexOptions.IgnoreCase);

            

        }
    }
}
