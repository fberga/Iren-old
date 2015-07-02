using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Forms;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    class Sheet : Base.Sheet
    {
        public Sheet(Excel.Worksheet ws) : base(ws)
        {

        }

        public override void CaricaInformazioni(bool all)
        {
            base.CaricaInformazioni(all);
            
            //profili PQNR
            if (_ws.Name == "Iren Termo")
            {
                DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
                DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
                DataView entitaRampa = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_RAMPA].DefaultView;
                categoriaEntita.RowFilter = "SiglaCategoria = '" + _siglaCategoria + "'";

                foreach (DataRowView entita in categoriaEntita)
                {
                    DateTime dataFine;
                    entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta LIKE '%GIORNI_struttura'";
                    
                    if (entitaProprieta.Count > 0)
                        dataFine = _dataInizio.AddDays(double.Parse("" + entitaProprieta[0]["Valore"]));
                    else
                        dataFine = _dataInizio.AddDays(Struct.intervalloGiorni);

                    double pRif =
                        (from r in DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA].AsEnumerable()
                         where r["SiglaEntita"].Equals(entita["SiglaEntita"])
                            && r["SiglaProprieta"].Equals("SISTEMA_COMANDI_PRIF")
                         select Double.Parse(r["Valore"].ToString())).FirstOrDefault();

                    int oreIntervallo = Date.GetOreIntervallo(dataFine);

                    Range rngPQNR = _definedNames.Get(entita["SiglaEntita"], "PQNR_PROFILO", Date.SuffissoDATA1).Extend(colOffset: oreIntervallo);

                    if (_ws.Range[rngPQNR.Columns[0].ToString()].Value != null)
                    {
                        int assetti = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_ASSETTO].AsEnumerable().Count(r => r["SiglaEntita"].Equals(entita["SiglaEntita"]));

                        double[] pMin = new double[oreIntervallo];
                        for (int i = 0; i < pMin.Length; i++) pMin[i] = double.MaxValue;

                        for (int i = 0; i < assetti; i++)
                        {
                            Range rngPmin = _definedNames.Get(entita["SiglaEntita"], "PMIN_TERNA_ASSETTO" + (i + 1), Date.SuffissoDATA1).Extend(colOffset: oreIntervallo);
                            for (int j = 0; j < oreIntervallo; j++)
                                pMin[j] = Math.Min(pMin[j], (double)(_ws.Range[rngPmin.Columns[j].ToString()].Value ?? 0d));
                        }

                        object[,] valori = new object[24, oreIntervallo];
                        for (int i = 0; i < oreIntervallo; i++)
                        {
                            pMin[i] = pMin[i] < pRif ? pRif : pMin[i];
                            entitaRampa.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaRampa = '" + _ws.Range[rngPQNR.Columns[i].ToString()].Value + "'";
                            if (entitaRampa.Count > 0)
                            {
                                for (int j = 0; j < 24; j++)
                                {
                                    if (entitaRampa[0]["Q" + (j + 1)] != DBNull.Value)
                                    {
                                        valori[j, i] = Math.Round(((int)entitaRampa[0]["Q" + (j + 1)]) * pRif / pMin[i]);
                                    }
                                }
                            }
                        }
                        Range rngPQNRVal = _definedNames.Get(entita["SiglaEntita"], "PQNR1", Date.SuffissoDATA1).Extend(rowOffset: 24, colOffset: oreIntervallo);
                        _ws.Range[rngPQNRVal.ToString()].Value = valori;
                    }
                }
            }
        }
    }
}
