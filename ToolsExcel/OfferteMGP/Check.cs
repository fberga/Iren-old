using Iren.ToolsExcel.Base;
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
                case 2:
                    n = CheckFunc2();
                    break;
                case 3:
                    n = CheckFunc3();
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

            DataView entitaParametroD = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PARAMETRO_D].DefaultView;
            entitaParametroD.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND SiglaParametro = 'LIMITE_PMAX'";
            double limitePmax = double.MaxValue;
            if (entitaParametroD.Count > 0)
                limitePmax = Double.Parse(entitaParametroD[0]["Valore"].ToString());

            entitaParametroD.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND SiglaParametro = 'LIMITE_PMIN'";
            double limitePmin = double.MinValue;
            if (entitaParametroD.Count > 0)
                limitePmin = Double.Parse(entitaParametroD[0]["Valore"].ToString());

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

                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "OFFERTA_MGP_E1", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double eOfferta1 = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "OFFERTA_MGP_E2", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double eOfferta2 = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "OFFERTA_MGP_E3", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double eOfferta3 = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "OFFERTA_MGP_E4", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double eOfferta4 = (double)(_ws.Range[rng.ToString()].Value ?? 0);

                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "OFFERTA_MGP_P1", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double pOfferta1 = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "OFFERTA_MGP_P2", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double pOfferta2 = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "OFFERTA_MGP_P3", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double pOfferta3 = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "OFFERTA_MGP_P4", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double pOfferta4 = (double)(_ws.Range[rng.ToString()].Value ?? 0);

                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "PCE", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double pce = (double)(_ws.Range[rng.ToString()].Value ?? 0);

                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "REQ", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double req = (double)(_ws.Range[rng.ToString()].Value ?? 0);

                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "MARGINEUP", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double margineUP = (double)(_ws.Range[rng.ToString()].Value ?? 0);

                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "PMIN", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double pmin = (double)(_ws.Range[rng.ToString()].Value ?? 0);

                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "PMAX", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double pmax = (double)(_ws.Range[rng.ToString()].Value ?? 0);

                TreeNode nOra = new TreeNode("Ora " + ora);

                bool errore = false;
                bool attenzione = false;

                if (eOfferta1 + eOfferta2 + eOfferta3 + eOfferta4 + pce > margineUP)
                {
                    nOra.Nodes.Add("Energia Offerta + PCE > Margine UP");
                    errore |= true;
                }
                if (eOfferta1 + eOfferta2 + eOfferta3 + eOfferta4 + pce > pmax)
                {
                    nOra.Nodes.Add("Energia Offerta + PCE > PMax");
                    errore |= true;
                }
                if (eOfferta1 + eOfferta2 + eOfferta3 + eOfferta4 + pce < pmin)
                {
                    nOra.Nodes.Add("Energia Offerta + PCE < PMin");
                    errore |= true;
                }
                if (eOfferta1 + eOfferta2 + eOfferta3 + eOfferta4 + pce < pmax)
                {
                    nOra.Nodes.Add("PCE > Preq");
                    attenzione |= true;
                }
                if (pce > pmax && pmax > 0)
                {
                    nOra.Nodes.Add("PCE > PMax");
                    errore |= true;
                }
                if (pce < 0)
                {
                    nOra.Nodes.Add("PCE < 0");
                    errore |= true;
                }
                if (pce > req)
                {
                    nOra.Nodes.Add("PCE > PReq");
                    errore |= true;
                }
                if (eOfferta1 == 0 && pOfferta1 != 0)
                {
                    nOra.Nodes.Add("Energia Offerta 1 = 0 e Prezzo Offerta 1 <> 0");
                    errore |= true;
                }
                if (eOfferta2 == 0 && pOfferta2 != 0)
                {
                    nOra.Nodes.Add("Energia Offerta 2 = 0 e Prezzo Offerta 2 <> 0");
                    errore |= true;
                }
                if (eOfferta3 == 0 && pOfferta3 != 0)
                {
                    nOra.Nodes.Add("Energia Offerta 3 = 0 e Prezzo Offerta 3 <> 0");
                    errore |= true;
                }
                if (eOfferta4 == 0 && pOfferta4 != 0)
                {
                    nOra.Nodes.Add("Energia Offerta 4 = 0 e Prezzo Offerta 4 <> 0");
                    errore |= true;
                }
                if (eOfferta1 != 0 && pOfferta1 == 0)
                {
                    nOra.Nodes.Add("Energia Offerta 1 <> 0 e Prezzo Offerta 1 = 0");
                    errore |= true;
                }
                if (eOfferta2 != 0 && pOfferta2 == 0)
                {
                    nOra.Nodes.Add("Energia Offerta 2 <> 0 e Prezzo Offerta 2 = 0");
                    errore |= true;
                }
                if (eOfferta3 != 0 && pOfferta3 == 0)
                {
                    nOra.Nodes.Add("Energia Offerta 3 <> 0 e Prezzo Offerta 3 = 0");
                    errore |= true;
                }
                if (eOfferta4 != 0 && pOfferta4 == 0)
                {
                    nOra.Nodes.Add("Energia Offerta 4 <> 0 e Prezzo Offerta 4 = 0");
                    errore |= true;
                }
                if (eOfferta1 + eOfferta2 + eOfferta3 + eOfferta4 + pce > limitePmax)
                {
                    nOra.Nodes.Add("Energia Offerta + PCE > Limite PMax");
                    errore |= true;
                }
                if (eOfferta1 + eOfferta2 + eOfferta3 + eOfferta4 + pce < limitePmin)
                {
                    nOra.Nodes.Add("Energia Offerta + PCE < Limite PMim");
                    errore |= true;
                }

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
        private CheckOutput CheckFunc2()
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

            DataView entitaParametroD = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PARAMETRO_D].DefaultView;
            entitaParametroD.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND SiglaParametro = 'LIMITE_PMAX'";
            double limitePmax = double.MaxValue;
            if (entitaParametroD.Count > 0)
                limitePmax = Double.Parse(entitaParametroD[0]["Valore"].ToString());

            entitaParametroD.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND SiglaParametro = 'LIMITE_PMIN'";
            double limitePmin = double.MinValue;
            if (entitaParametroD.Count > 0)
                limitePmin = Double.Parse(entitaParametroD[0]["Valore"].ToString());

            for (int i = 1; i <= rngCheck.ColOffset; i++)
            {
                string suffissoData = Utility.Date.GetSuffissoData(Utility.DataBase.DataAttiva.AddHours(i - 1));
                if (data != Utility.DataBase.DataAttiva.AddHours(i - 1).ToString("dd-MM-yyyy"))
                {
                    data = Utility.DataBase.DataAttiva.AddHours(i - 1).ToString("dd-MM-yyyy");
                    if (nData.Nodes.Count > 0)
                        n.Nodes.Add(nData);

                    nData = new TreeNode(data);
                }

                int ora = (i - 1) % Utility.Date.GetOreGiorno(suffissoData) + 1;

                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "OFFERTA_MGP_E1", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double eOfferta1 = (double)(_ws.Range[rng.ToString()].Value ?? 0);
               
                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "PCE", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double pce = (double)(_ws.Range[rng.ToString()].Value ?? 0);

                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "PROGR_UC", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double progrUC = (double)(_ws.Range[rng.ToString()].Value ?? 0);

                bool errore = false;
                bool attenzione = false;

                TreeNode nOra = new TreeNode("Ora " + ora);

                if (eOfferta1 + pce != progrUC)
                {
                    nOra.Nodes.Add("Eofferta + PCE <> Programma");
                    errore |= true;
                }
                if (eOfferta1 + pce > limitePmax)
                {
                    nOra.Nodes.Add("Eofferta + PCE > PLimMax");
                    errore |= true;
                }
                if (eOfferta1 + pce < limitePmin)
                {
                    nOra.Nodes.Add("Eofferta + PCE < PLimMin");
                    errore |= true;
                }
                if (progrUC < pce)
                {
                    nOra.Nodes.Add("PCE > Programma");
                    attenzione |= true;
                }

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
        private CheckOutput CheckFunc3()
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

            DataView entitaParametroD = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PARAMETRO_D].DefaultView;
            entitaParametroD.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND SiglaParametro = 'LIMITE_PMAX'";
            double limitePmax = double.MaxValue;
            if (entitaParametroD.Count > 0)
                limitePmax = Double.Parse(entitaParametroD[0]["Valore"].ToString());

            entitaParametroD.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND SiglaParametro = 'LIMITE_PMIN'";
            double limitePmin = double.MinValue;
            if (entitaParametroD.Count > 0)
                limitePmin = Double.Parse(entitaParametroD[0]["Valore"].ToString());

            for (int i = 1; i <= rngCheck.ColOffset; i++)
            {
                string suffissoData = Utility.Date.GetSuffissoData(Utility.DataBase.DataAttiva.AddHours(i - 1));
                if (data != Utility.DataBase.DataAttiva.AddHours(i - 1).ToString("dd-MM-yyyy"))
                {
                    data = Utility.DataBase.DataAttiva.AddHours(i - 1).ToString("dd-MM-yyyy");
                    if (nData.Nodes.Count > 0)
                        n.Nodes.Add(nData);

                    nData = new TreeNode(data);
                }

                int ora = (i - 1) % Utility.Date.GetOreGiorno(suffissoData) + 1;

                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "OFFERTA_MGP_E1", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double eOfferta1 = (double)(_ws.Range[rng.ToString()].Value ?? 0);

                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "PCE", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double pce = (double)(_ws.Range[rng.ToString()].Value ?? 0);

                rng = _newNomiDefiniti.Get(_check.SiglaEntita, "PROGR_UC", suffissoData, Utility.Date.GetSuffissoOra(ora));
                double progrUC = (double)(_ws.Range[rng.ToString()].Value ?? 0);

                double delta = 0;
                if (_newNomiDefiniti.IsDefined(_check.SiglaEntita, "DELTA_PROGR_UC"))
                {
                    rng = _newNomiDefiniti.Get(_check.SiglaEntita, "DELTA_PROGR_UC", suffissoData, Utility.Date.GetSuffissoOra(ora));
                    delta = (double)(_ws.Range[rng.ToString()].Value ?? 0);
                }

                bool errore = false;
                bool attenzione = false;

                TreeNode nOra = new TreeNode("Ora " + ora);

                if (eOfferta1 + pce != progrUC)
                {
                    nOra.Nodes.Add("Eofferta + PCE <> Programma");
                    errore |= true;
                }
                if (eOfferta1 + pce > limitePmax)
                {
                    nOra.Nodes.Add("Eofferta + PCE > PLimMax");
                    errore |= true;
                }
                if (eOfferta1 + pce < limitePmin)
                {
                    nOra.Nodes.Add("Eofferta + PCE < PLimMin");
                    errore |= true;
                }
                if (progrUC + delta < pce)
                {
                    nOra.Nodes.Add("PCE > Programma + Delta");
                    attenzione |= true;
                }

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
