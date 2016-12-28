using System;
using System.Configuration;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace GeneraXls
{
    static class Program
    {
        /// <summary>
        /// Punto di ingresso principale dell'applicazione.
        /// </summary>
        [STAThread]
        static void Main()
        {
            string tipoMercato;
            string pathInput;
            string pathOutput;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FrmGeneraXls());

            tipoMercato = ConfigurationManager.AppSettings.Get("tipoMercato");
            pathInput = ConfigurationManager.AppSettings.Get("pathInput");
            pathOutput = ConfigurationManager.AppSettings.Get("pathOutput");


        }
    }
}
