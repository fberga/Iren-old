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
        public override CheckOutput ExecuteCheck(Excel.Worksheet ws, DefinedNames definedNames, CheckObj check)
        {
            _ws = ws;
            _nomiDefiniti = definedNames;
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
                decimal ePPA = GetDecimal(_check.SiglaEntita, "PEM", suffissoData, Utility.Date.GetSuffissoOra(ora));
                decimal ePSMaxAccettata = GetDecimal(_check.SiglaEntita, "PSMAX_ACCETTATA", suffissoData, Utility.Date.GetSuffissoOra(ora));
                decimal ePSMinAccettata = GetDecimal(_check.SiglaEntita, "PSMIN_ACCETTATA", suffissoData, Utility.Date.GetSuffissoOra(ora));
                //fine caricameto dati

                TreeNode nOra = new TreeNode("Ora " + ora);

                bool errore = false;
                bool attenzione = false;

                //controlli
                if (ePPA > ePSMaxAccettata)
                {
                    nOra.Nodes.Add("Programma di produzione superiore alla PMax Terna");
                    errore |= true;
                }
                if (ePSMaxAccettata < ePPA && ePPA < ePSMinAccettata && ePPA > 0)
                {
                    nOra.Nodes.Add("Programma di produzione non coerente con PMin-PMax Terna");
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
