using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Funzione di check.
    /// </summary>
    class Check : Base.Check
    {
        public override CheckOutput ExecuteCheck(Excel.Worksheet ws, DefinedNames definedNames, CheckObj check)
        {
            _ws = ws;
            _nomiDefiniti = definedNames;
            _check = check;

            return CheckFunc1();
        }

        private CheckOutput CheckFunc1() 
        {
            Range rngCheck = new Range(_check.Range);

            DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + _check.SiglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;

            TreeNode n = new TreeNode(categoriaEntita[0]["DesEntita"].ToString());
            n.Name = _check.SiglaEntita;
            TreeNode nData = new TreeNode();
            string data = "";

            CheckOutput.CheckStatus status = CheckOutput.CheckStatus.Ok;

            var assettiFasce = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].AsEnumerable()
                .Where(r => r["SiglaEntita"].Equals(_check.SiglaEntita) && r["IdApplicazione"].Equals(Workbook.IdApplicazione))
                .Where(r => r["SiglaInformazione"].ToString().StartsWith("PSMAX_ASSETTO") && r["Visibile"].Equals("1"))
                .Select(r => r["SiglaInformazione"].ToString().Replace("PSMAX_", ""));

            for (int i = 1; i <= rngCheck.ColOffset; i++)
            {
                string suffissoData = Date.GetSuffissoData(DataBase.DataAttiva.AddHours(i - 1));
                if (data != DataBase.DataAttiva.AddHours(i - 1).ToString("dd-MM-yyyy"))
                {
                    data = DataBase.DataAttiva.AddHours(i - 1).ToString("dd-MM-yyyy");
                    if(nData.Nodes.Count > 0)
                        n.Nodes.Add(nData);

                    nData = new TreeNode(data);
                }

                int ora = (i - 1) % Date.GetOreGiorno(suffissoData) + 1;

                //caricamento dati                
                object profiloPQNR = GetObject(_check.SiglaEntita, "PQNR_PROFILO", suffissoData, Date.GetSuffissoOra(ora));
                //fine caricameto dati

                TreeNode nOra = new TreeNode("Ora " + ora);

                bool errore = false;
                bool attenzione = false;

                foreach (var assettoFascia in assettiFasce)
                {
                    //caricamento dati
                    decimal psmax = GetDecimal(_check.SiglaEntita, "PSMAX_CORRETTA_" + assettoFascia, suffissoData, Date.GetSuffissoOra(ora));
                    if (psmax == 0)
                        psmax = GetDecimal(_check.SiglaEntita, "PSMAX_" + assettoFascia, suffissoData, Date.GetSuffissoOra(ora));

                    decimal psmaxQ1 = GetDecimal(_check.SiglaEntita, "PSMAXQ1_" + assettoFascia, suffissoData, Date.GetSuffissoOra(ora));
                    decimal psmaxQ2 = GetDecimal(_check.SiglaEntita, "PSMAXQ2_" + assettoFascia, suffissoData, Date.GetSuffissoOra(ora));
                    decimal psmaxQ3 = GetDecimal(_check.SiglaEntita, "PSMAXQ3_" + assettoFascia, suffissoData, Date.GetSuffissoOra(ora));
                    decimal psmaxQ4 = GetDecimal(_check.SiglaEntita, "PSMAXQ4_" + assettoFascia, suffissoData, Date.GetSuffissoOra(ora));


                    decimal psmin = GetDecimal(_check.SiglaEntita, "PSMIN_CORRETTA_" + assettoFascia, suffissoData, Date.GetSuffissoOra(ora));
                    if (psmin == 0)
                        psmin = GetDecimal(_check.SiglaEntita, "PSMIN_" + assettoFascia, suffissoData, Date.GetSuffissoOra(ora));

                    decimal psminQ1 = GetDecimal(_check.SiglaEntita, "PSMINQ1_" + assettoFascia, suffissoData, Date.GetSuffissoOra(ora));
                    decimal psminQ2 = GetDecimal(_check.SiglaEntita, "PSMINQ2_" + assettoFascia, suffissoData, Date.GetSuffissoOra(ora));
                    decimal psminQ3 = GetDecimal(_check.SiglaEntita, "PSMINQ3_" + assettoFascia, suffissoData, Date.GetSuffissoOra(ora));
                    decimal psminQ4 = GetDecimal(_check.SiglaEntita, "PSMINQ4_" + assettoFascia, suffissoData, Date.GetSuffissoOra(ora));
                    //fine caricameto dati

                    //controlli
                    if (psmax != psmaxQ1)
                    {
                        nOra.Nodes.Add("PSMAX accettata 0-15 <> PSMAX");
                        attenzione |= true;
                    }
                    if (psmax != psmaxQ2)
                    {
                        nOra.Nodes.Add("PSMAX accettata 15-30 <> PSMAX");
                        attenzione |= true;
                    }
                    if (psmax != psmaxQ3)
                    {
                        nOra.Nodes.Add("PSMAX accettata 30-45 <> PSMAX");
                        attenzione |= true;
                    }
                    if (psmax != psmaxQ4)
                    {
                        nOra.Nodes.Add("PSMAX accettata 45-60 <> PSMAX");
                        attenzione |= true;
                    }
                    /////////////////////////////////////////////////////////////
                    if (psmin != psminQ1)
                    {
                        nOra.Nodes.Add("PSMIN accettata 0-15 <> PSMIN");
                        attenzione |= true;
                    }
                    if (psmin != psminQ2)
                    {
                        nOra.Nodes.Add("PSMIN accettata 15-30 <> PSMIN");
                        attenzione |= true;
                    }
                    if (psmin != psminQ3)
                    {
                        nOra.Nodes.Add("PSMIN accettata 30-45 <> PSMIN");
                        attenzione |= true;
                    }
                    if (psmin != psminQ4)
                    {
                        nOra.Nodes.Add("PSMIN accettata 45-60 <> PSMIN");
                        attenzione |= true;
                    }
                    //fine controlli
                }

                //controlli
                if (profiloPQNR == null)
                {
                    nOra.Nodes.Add("Il profilo PQNR non è selezionato");
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
