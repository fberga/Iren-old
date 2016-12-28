using Iren.PSO.Base;
using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Deployment.Application;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Mono.Options;
using System.Collections.Generic;
using System.Xml.Linq;

namespace Iren.PSO.ConsoleLauncher
{
    class Program
    {
        private static Excel.Application _xlApp;

        static void Main(string[] args)
        {
            // these variables will be set when the command line is parsed
            int idApplicazione = -1;
            bool accettaCambioData = false;
            bool rifiutaCambioData = false;
            bool aggiornaStruttura = false;
            bool aggiornaDati = false;
            bool carica = false;
            bool genera = false;
            bool esporta = false;
            bool shouldShowHelp = false;

            // thses are the available options, not that they set the variables
            OptionSet options = new OptionSet { 
                { "i|idApp=", "l'id dell'applicazione.", (int id) => idApplicazione = id }, 
                { "a|accdate", "accetta automaticamente il cambio data", cd => accettaCambioData = cd != null }, 
                { "r|rifdate", "rifiuta automaticamente il cambio data", cd => rifiutaCambioData = cd != null },
                { "s|aggstr", "forza l'aggiornamento della struttura e dati", aggstr => aggiornaStruttura = aggstr != null},
                { "d|aggdati", "forza l'aggiornamento dei dati ma non della struttura", aggdt => aggiornaDati = aggdt != null},
                { "c|carica", "esegue l'azione Carica", c => carica = c != null},
                { "g|genera", "esegue l'azione genera", g => genera = g != null},
                { "e|esporta", "esegue l'azione Esporta", e => esporta = e != null},
                { "h|help", "mostra questo help ed esce", h => shouldShowHelp = h != null }
            };


            List<string> extra;
            try
            {
                // parse the command line
                extra = options.Parse(args);
            }
            catch (OptionException e)
            {
                // output some error message
                Console.Write("ConsoleLauncher: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `consolelauncher --help' for more information.");
                return;
            }

            if (shouldShowHelp)
            {
                // show some app description message
                Console.WriteLine("Utilizzo: ConsoleLauncer.exe [OPTIONS]+");
                Console.WriteLine("Esegue un'applicazione della suite PSO dal terminale.");
                Console.WriteLine("L'IdApplicazione deve essere specificato.");
                Console.WriteLine();

                // output the options
                Console.WriteLine("Options:");
                options.WriteOptionDescriptions(Console.Out);
                return;
            }


            Excel.Workbooks wbs = null;
            try
            {
                wbs = _xlApp.Workbooks;
            }
            catch
            {
                _xlApp = new Excel.Application();
            }
            finally
            {
                if (wbs != null) Marshal.ReleaseComObject(wbs);
                wbs = null;
            }

            XDocument doc = new XDocument(new XDeclaration("1.0", "ISO-8859-1", "yes"),
                new XElement("AvvioAutomatico",
                new XElement("AccettaCambioData", accettaCambioData),
                new XElement("RifiutaCambioData", rifiutaCambioData),
                new XElement("AggiornaStruttura", aggiornaStruttura),
                new XElement("AggiornaDati", aggiornaDati),
                new XElement("Carica", carica),
                new XElement("Genera", genera),
                new XElement("Esporta", esporta)));

            doc.Save(@"C:\Emergenza\AvvioAutomatico.xml");

            Workbook.AvviaApplicazione(_xlApp, idApplicazione);
        }
    }
}
