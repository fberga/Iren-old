﻿using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Funzioni di caricamento personalizzato. Una volta caricati i dati, scrive l'informazione anche nei fogli di export.
    /// </summary>
    public class Carica : Base.Carica
    {
        //public override bool AzioneInformazione(object siglaEntita, object siglaAzione, object azionePadre, DateTime giorno, object parametro = null)
        //{
        //    bool o = base.AzioneInformazione(siglaEntita, siglaAzione, azionePadre, giorno, parametro);

        //    Riepilogo main = new Riepilogo();
        //    main.AggiornaPrevisione(siglaEntita);

        //    return o;
        //}
    }
}
