﻿using Iren.ToolsExcel.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    class Check : Base.Check
    {
        Excel.Worksheet _ws;
        NewDefinedNames _newNomiDefiniti;
        CheckObj _check;

        public override CheckOutput ExecuteCheck(Excel.Worksheet ws, NewDefinedNames newNomiDefiniti, CheckObj check)
        {
            _ws = ws;
            _newNomiDefiniti = newNomiDefiniti;
            _check = check;

            CheckOutput n = new CheckOutput();

            switch (check.Type)
            {
                case 1:
                    n = CheckFunc1();
                    break;               
            }

            return n;
        }

        private CheckOutput CheckFunc1()
        {
            Range rngCheck = new Range(_check.Range);
            Range rng;

            DataView categoriaEntita = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "'";

            TreeNode n = new TreeNode(categoriaEntita[0]["DesEntita"].ToString());
            n.Name = _check.SiglaEntita;
            TreeNode nData = new TreeNode();
            string data = "";

            CheckOutput.CheckStatus status = CheckOutput.CheckStatus.Ok;

            for (int i = 1; i <= rngCheck.ColOffset; i++)
            {
                string suffissoData = Utility.Date.GetSuffissoData(Utility.DataBase.DataAttiva.AddHours(i - 1));
                if (data != Utility.DataBase.DataAttiva.AddHours(i - 1).ToString("dd-MM-yyyy"))
                {
                    data = Utility.DataBase.DataAttiva.AddHours(i - 1).ToString("dd-MM-yyyy");
                    if(nData.Nodes.Count > 0)
                        n.Nodes.Add(nData);

                    nData = new TreeNode(data);
                }

                int ora = (i - 1) % Utility.Date.GetOreGiorno(suffissoData) + 1;

                //caricamento dati
                rng = _newNomiDefiniti.Get("CE_MTX", "TEMPERATURA", suffissoData, Utility.Date.GetSuffissoOra(ora));
                object temperaturaMTX = _ws.Range[rng.ToString()].Value;
                rng = _newNomiDefiniti.Get("CE_TTX", "TEMPERATURA", suffissoData, Utility.Date.GetSuffissoOra(ora));
                object temperaturaTTX = _ws.Range[rng.ToString()].Value;
                rng = _newNomiDefiniti.Get("CE_MTX", "PRESSIONE", suffissoData, Utility.Date.GetSuffissoOra(ora));
                object pressioneMTX = _ws.Range[rng.ToString()].Value;
                rng = _newNomiDefiniti.Get("CE_TTX", "PRESSIONE", suffissoData, Utility.Date.GetSuffissoOra(ora));
                object pressioneTTX = _ws.Range[rng.ToString()].Value;
                rng = _newNomiDefiniti.Get("CT_TORINO", "CARICO_TERMICO_PREVISIONE", suffissoData, Utility.Date.GetSuffissoOra(ora));
                object caricoTermico = _ws.Range[rng.ToString()].Value;
                rng = _newNomiDefiniti.Get("ZM_NORD", "PREV_PREZZO", suffissoData, Utility.Date.GetSuffissoOra(ora));
                object prezzoZonale = _ws.Range[rng.ToString()].Value;
                rng = _newNomiDefiniti.Get("CE_MTX", "PREV_PORTATA", suffissoData, Utility.Date.GetSuffissoOra(ora));
                object portataCanale = _ws.Range[rng.ToString()].Value;
                rng = _newNomiDefiniti.Get("CE_TTX", "GRUPPO_FRIGO", suffissoData, Utility.Date.GetSuffissoOra(ora));
                object gruppoFrigo = _ws.Range[rng.ToString()].Value;
                rng = _newNomiDefiniti.Get("UP_MT2R", "TEMP_PROG1", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double onOff = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                rng = _newNomiDefiniti.Get("UP_MT3", "TEMP_PROG1", suffissoData, Utility.Date.GetSuffissoOra(ora));
                onOff += (double)(_ws.Range[rng.ToString()].Value ?? 0);
                rng = _newNomiDefiniti.Get("UP_MT2R", "UNIT_COMM", suffissoData, Utility.Date.GetSuffissoOra(ora));
                string unitCommMT2R = (string)(_ws.Range[rng.ToString()].Value ?? "");
                rng = _newNomiDefiniti.Get("UP_MT3", "UNIT_COMM", suffissoData, Utility.Date.GetSuffissoOra(ora));
                string unitCommMT3 = (string)(_ws.Range[rng.ToString()].Value ?? "");
                rng = _newNomiDefiniti.Get("UP_TN1", "UNIT_COMM", suffissoData, Utility.Date.GetSuffissoOra(ora));
                string unitCommTN1 = (string)(_ws.Range[rng.ToString()].Value ?? "");
                rng = _newNomiDefiniti.Get("UP_MT2R", "DISPONIBILITA_CALORE_PMAX", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double dispCalorePMaxMT2R = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                rng = _newNomiDefiniti.Get("UP_MT3", "DISPONIBILITA_CALORE_PMAX", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double dispCalorePMaxMT3 = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                rng = _newNomiDefiniti.Get("UP_TN1", "DISPONIBILITA_CALORE_PMAX", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double dispCalorePMaxTN1 = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                rng = _newNomiDefiniti.Get("UP_MT2R", "DISPONIBILITA_CALORE_PMIN", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double dispCalorePMinMT2R = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                rng = _newNomiDefiniti.Get("UP_MT3", "DISPONIBILITA_CALORE_PMIN", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double dispCalorePMinMT3 = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                rng = _newNomiDefiniti.Get("UP_TN1", "DISPONIBILITA_CALORE_PMIN", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double dispCalorePMinTN1 = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                //fine caricameto dati

                TreeNode nOra = new TreeNode("Ora " + ora);

                bool errore = false;
                bool attenzione = false;

                //controlli
                if (temperaturaMTX == null)
                {
                    nOra.Nodes.Add("Temperatura Moncalieri assente");
                    errore |= true;
                }
                if (temperaturaMTX != null && (double)temperaturaMTX < -20)
                {
                    nOra.Nodes.Add("Temperatura Moncalieri < soglia minima");
                    errore |= true;
                }
                if (temperaturaMTX != null && (double)temperaturaMTX > 45)
                {
                    nOra.Nodes.Add("Temperatura Moncalieri > soglia massima");
                    errore |= true;
                }
                if (temperaturaTTX == null)
                {
                    nOra.Nodes.Add("Temperatura Torino Nord assente");
                    errore |= true;
                }
                if (temperaturaTTX != null && (double)temperaturaTTX < -20)
                {
                    nOra.Nodes.Add("Temperatura Torino Nord < soglia minima");
                    errore |= true;
                }
                if (temperaturaTTX != null && (double)temperaturaTTX > 45)
                {
                    nOra.Nodes.Add("Temperatura Torino Nord > soglia massima");
                    errore |= true;
                }
                if (pressioneMTX == null)
                {
                    nOra.Nodes.Add("Pressione Moncalieri assente");
                    errore |= true;
                }
                if (pressioneMTX != null && (double)pressioneMTX < -850)
                {
                    nOra.Nodes.Add("Pressione Moncalieri < soglia minima");
                    errore |= true;
                }
                if (pressioneMTX != null && (double)pressioneMTX > 1100)
                {
                    nOra.Nodes.Add("Pressione Moncalieri > soglia massima");
                    errore |= true;
                }
                if (pressioneTTX == null)
                {
                    nOra.Nodes.Add("Pressione Torino Nord assente");
                    errore |= true;
                }
                if (pressioneTTX != null && (double)pressioneTTX < -850)
                {
                    nOra.Nodes.Add("Pressione Torino Nord < soglia minima");
                    errore |= true;
                }
                if (pressioneTTX != null && (double)pressioneTTX > 1100)
                {
                    nOra.Nodes.Add("Pressione Torino Nord > soglia massima");
                    errore |= true;
                }
                if (caricoTermico == null)
                {
                    nOra.Nodes.Add("Carico termico assente");
                    errore |= true;
                }
                if (caricoTermico != null && (double)caricoTermico < 10)
                {
                    nOra.Nodes.Add("Carico termico < soglia minima");
                    errore |= true;
                }
                if (caricoTermico != null && (double)caricoTermico > 2000)
                {
                    nOra.Nodes.Add("Carico termico > soglia massima");
                    errore |= true;
                }
                if (prezzoZonale == null)
                {
                    nOra.Nodes.Add("Prezzo zonale assente");
                    errore |= true;
                }
                if (prezzoZonale != null && (double)prezzoZonale < 0)
                {
                    nOra.Nodes.Add("Prezzo zonale < soglia minima");
                    errore |= true;
                }
                if (prezzoZonale != null && (double)prezzoZonale > 500)
                {
                    nOra.Nodes.Add("Prezzo zonale > soglia massima");
                    errore |= true;
                }
                if (portataCanale == null)
                {
                    nOra.Nodes.Add("Portata canale assente");
                    errore |= true;
                }
                if(portataCanale != null && (
                        ((double)portataCanale < 7 && (unitCommMT2R.Equals("off") || unitCommMT2R.Equals("m") || unitCommMT3.Equals("off") || unitCommMT3.Equals("m"))) 
                     || ((double)portataCanale < 14 && ((unitCommMT2R.Equals("off") || unitCommMT2R.Equals("m")) && (unitCommMT3.Equals("off") || unitCommMT3.Equals("m"))))))
                {
                    nOra.Nodes.Add("Portata canale < soglia minima");
                    errore |= true;
                }
                if (portataCanale != null && (double)portataCanale > 90)
                {
                    nOra.Nodes.Add("Portata canale > soglia massima");
                    errore |= true;
                }
                if (gruppoFrigo == null)
                {
                    nOra.Nodes.Add("Numero gruppi frigo assente");
                    errore |= true;
                }
                if (gruppoFrigo != null && (double)gruppoFrigo < 0)
                {
                    nOra.Nodes.Add("Numero gruppi frigo < soglia minima");
                    errore |= true;
                }
                if (gruppoFrigo != null && (double)gruppoFrigo > 6)
                {
                    nOra.Nodes.Add("Numero gruppi frigo > soglia massima");
                    errore |= true;
                }
                if(dispCalorePMinMT2R > 0 && (unitCommMT2R == "ind" || unitCommMT2R == "off"))
                {
                    nOra.Nodes.Add("MT2R disponibilità minima calore > 0");
                    errore |= true;
                }
                if (dispCalorePMinMT3 > 0 && (unitCommMT3 == "ind" || unitCommMT3 == "off"))
                {
                    nOra.Nodes.Add("MT3 disponibilità minima calore > 0");
                    errore |= true;
                }
                if (dispCalorePMinTN1 > 0 && (unitCommTN1 == "ind" || unitCommTN1 == "off"))
                {
                    nOra.Nodes.Add("TN1 disponibilità minima calore > 0");
                    errore |= true;
                }
                if (dispCalorePMinMT2R > dispCalorePMaxMT2R)
                {
                    nOra.Nodes.Add("MT2R disponibilità minima calore > disponibilità massima calore");
                    errore |= true;
                }
                if (dispCalorePMinMT3 > dispCalorePMaxMT3)
                {
                    nOra.Nodes.Add("MT3 disponibilità minima calore > disponibilità massima calore");
                    errore |= true;
                }
                if (dispCalorePMinTN1 > dispCalorePMaxTN1)
                {
                    nOra.Nodes.Add("TN1 disponibilità minima calore > disponibilità massima calore");
                    errore |= true;
                }
                //fine controlli

                if (errore)
                {
                    ErrorStyle(ref nOra);
                    status = CheckOutput.CheckStatus.Error;
                }
                else if (attenzione)
                {
                    AlertStyle(ref nOra);
                    if (status != CheckOutput.CheckStatus.Error)
                        status = CheckOutput.CheckStatus.Alert;
                }

                nOra.Name = "'" + _ws.Name + "'!" + rngCheck.Columns[i - 1].ToString();

                if (nOra.Nodes.Count > 0)
                    nData.Nodes.Add(nOra);

                string value = errore ? "ERRORE" : attenzione ? "ATTENZ." : "OK";
                _ws.Range[rngCheck.Columns[i - 1].ToString()].Value = value;
            }
            
            if (nData.Nodes.Count > 0)
            {
                n.Nodes.Add(nData);
            }

            if (n.Nodes.Count > 0)
                return new CheckOutput(n, status);

            return new CheckOutput();
        }
    }
}
