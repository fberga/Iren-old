﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Data;
using System.Globalization;

namespace Iren.FrontOffice.Base
{
    public class Riepilogo<T>: CommonFunctions
    {
        #region Variabili
        
        Worksheet _ws;
        Dictionary<string, object> _config;
        DateTime _dataInizio;
        DateTime _dataFine;
        DefinedNames _nomiDefiniti;
        Cell _cell;
        Struttura _struttura;
        int _rigaAttiva;
        int _colonnaInizio;


        #endregion

        #region Costruttori

        public Riepilogo(T categoria)
        {
            Type t = categoria.GetType();
            PropertyInfo p = t.GetProperty("Base");
            _ws = (Worksheet)p.GetValue(categoria, null);

            FieldInfo f = t.GetField("config");
            _config = (Dictionary<string, object>)f.GetValue(categoria);

            //dimensionamento celle in base ai parametri del DB
            DataView paramApplicazione = LocalDB.Tables[Tab.APPLICAZIONE].DefaultView;

            _cell = new Cell();
            _struttura = new Struttura();

            //prendo i valori di default
            _cell.Width.empty = double.Parse(paramApplicazione[0]["ColVuotaWidth"].ToString());
            _cell.Width.dato = double.Parse(paramApplicazione[0]["ColDatoWidth"].ToString());
            _cell.Width.entita = double.Parse(paramApplicazione[0]["ColEntitaWidth"].ToString());
            _cell.Width.informazione = double.Parse(paramApplicazione[0]["ColInformazioneWidth"].ToString());
            _cell.Width.unitaMisura = double.Parse(paramApplicazione[0]["ColUMWidth"].ToString());
            _cell.Width.parametro = double.Parse(paramApplicazione[0]["ColParametroWidth"].ToString());
            _cell.Height.normal = double.Parse(paramApplicazione[0]["RowHeight"].ToString());
            _cell.Height.empty = double.Parse(paramApplicazione[0]["RowVuotaHeight"].ToString());
            
            _struttura.rigaBlock = 5;
            _struttura.intervalloGiorni = (int)paramApplicazione[0]["IntervalloGiorni"];
            _struttura.colBlock = 59;

            _nomiDefiniti = new DefinedNames(_ws.Name);
        }

        #endregion

        #region Metodi

        private void CicloGiorni(Func<int, string, DateTime, bool> callback)
        {
            for (DateTime giorno = _dataInizio; giorno <= _dataFine; giorno = giorno.AddDays(1))
            {
                int oreGiorno = GetOreGiorno(giorno);
                string suffissoData = GetSuffissoData(_dataInizio, giorno);

                if (giorno == _dataInizio && _struttura.visData0H24)
                {
                    oreGiorno++;
                }

                callback(oreGiorno, suffissoData, giorno);
            }
        }

        private void Clear()
        {
            int dataOreTot = GetOreIntervallo(_dataInizio, _dataInizio.AddDays(_struttura.intervalloGiorni)) + (_struttura.visData0H24 ? 1 : 0) + (_struttura.visParametro ? 1 : 0);

            _ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;

            _ws.UsedRange.EntireColumn.Delete();
            _ws.UsedRange.FormatConditions.Delete();
            _ws.UsedRange.EntireRow.Hidden = false;
            _ws.UsedRange.Font.Size = 8;
            _ws.UsedRange.NumberFormat = "General";
            _ws.UsedRange.Font.Name = "Verdana";

            _ws.Range[_ws.Cells[1, 1], _ws.Cells[1, _struttura.colRecap - 1]].EntireColumn.ColumnWidth = _cell.Width.empty;
            _ws.Rows[_struttura.rigaGoto].RowHeight = _cell.Height.normal;
            _ws.Rows[1].RowHeight = _cell.Height.empty;

            _ws.Activate();
            _ws.Application.ActiveWindow.FreezePanes = false;
            _ws.Cells[_struttura.rigaBlock, _struttura.colBlock].Select();
            _ws.Application.ActiveWindow.ScrollColumn = 1;
            _ws.Application.ActiveWindow.ScrollRow = 1;
            _ws.Application.ActiveWindow.FreezePanes = true;
        }

        public void LoadStructure()
        {
            Clear();

            _colonnaInizio = _struttura.colRecap;
            _rigaAttiva = _struttura.rowRecap;

            InitBarraTitolo();
        }

        private void InitBarraTitolo()
        {
            DataView azioni = LocalDB.Tables[Tab.AZIONE].DefaultView;

            azioni.RowFilter = "GERARCHIA IS NOT NULL";
            int count = azioni.Count;
            azioni.RowFilter = "";

            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                Excel.Range rng = _ws.Range[_ws.Cells[_rigaAttiva, _colonnaInizio], _ws.Cells[_rigaAttiva + 2, _colonnaInizio + count - 1]];
                rng.Style = "recapTitleBarStyle";
                rng.Rows[1].Merge();

                return true;
            });
        }

        #endregion

    }
}
