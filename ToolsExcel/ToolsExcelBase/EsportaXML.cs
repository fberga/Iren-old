using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Base
{
    public class EsportaXML
    {
        public EsportaXML()
        {
        }


        public void RunExport()
        {
            DataTable categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA];
            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            DataView entitaInformazione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;

            foreach (DataRow entita in categoriaEntita.Rows)
            {
                object siglaEntita = entita["SiglaEntita"];
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                if (nomeFoglio != "")
                {
                    DefinedNames definedNames = new DefinedNames(nomeFoglio);

                    Excel.Worksheet ws = Workbook.Sheets[nomeFoglio];

                    bool hasData0H24 = definedNames.HasData0H24;

                    entitaInformazione.RowFilter = "(SiglaEntita = '" + siglaEntita + "' OR SiglaEntitaRif = '" + siglaEntita + "') AND Editabile = '1' AND IdApplicazione = " + Simboli.AppID;

                    DataTable entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA];
                    int intervalloGiorni =
                        (from r in entitaProprieta.AsEnumerable()
                         where r["IdApplicazione"].Equals(int.Parse(Simboli.AppID)) && r["SiglaEntita"].Equals(siglaEntita) && r["SiglaProprieta"].ToString().EndsWith("GIORNI_STRUTTURA")
                         select int.Parse(r["Valore"].ToString())).FirstOrDefault();

                    DateTime dataFine = DataBase.DataAttiva.AddDays(Math.Max(
                        (from r in entitaProprieta.AsEnumerable()
                         where r["IdApplicazione"].Equals(int.Parse(Simboli.AppID)) && r["SiglaEntita"].Equals(siglaEntita) && r["SiglaProprieta"].ToString().EndsWith("GIORNI_STRUTTURA")
                         select int.Parse(r["Valore"].ToString())).FirstOrDefault(), Struct.intervalloGiorni));

                    foreach (DataRowView info in entitaInformazione)
                    {
                        object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];

                        if (Struct.tipoVisualizzazione == "O")
                        {
                            //prima cella della riga da salvare (non considera Data0H24)
                            Range rng = definedNames.Get(siglaEntitaRif, info["SiglaInformazione"], Date.SuffissoDATA1).Extend(colOffset: Date.GetOreIntervallo(dataFine));
                            Handler.StoreEdit(ws.Range[rng.ToString()], 0, true, DataBase.Tab.EXPORT_XML);
                        }
                        else
                        {
                            for (DateTime giorno = DataBase.DataAttiva; giorno <= dataFine; giorno = giorno.AddDays(1))
                            {
                                Range rng = definedNames.Get(siglaEntitaRif, info["SiglaInformazione"], Date.GetSuffissoData(giorno)).Extend(colOffset: Date.GetOreGiorno(giorno));
                                Handler.StoreEdit(ws.Range[rng.ToString()], 0, true, DataBase.Tab.EXPORT_XML);
                            }
                        }
                    }
                }
            }

            //preparo l'export
            var path = Workbook.GetUsrConfigElement("emergenza");
            //path della cartella di emergenza
            string cartellaEmergenza = Esporta.PreparePath(path.Value);
            string cartellaExport = Path.Combine(cartellaEmergenza, Simboli.nomeApplicazione.Replace(" ", ""));
            string fileName = Path.Combine(cartellaExport, Simboli.nomeApplicazione.Replace(" ", "").ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml");

            if (!Directory.Exists(cartellaExport))
                Directory.CreateDirectory(cartellaExport);

            DataTable export = DataBase.LocalDB.Tables[DataBase.Tab.EXPORT_XML];
            export.WriteXml(fileName);

            //svuoto la tabella alla fine dell'utilizzo
            export.Clear();
        }
    }
}
