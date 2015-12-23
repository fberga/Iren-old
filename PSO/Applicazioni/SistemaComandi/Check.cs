﻿using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
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
