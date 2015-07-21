using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    public class Aggiorna : Base.Aggiorna
    {
        public Aggiorna()
            : base()
        {

        }

        public override bool Struttura()
        {
            //Aggiungo i fogli dei mercati leggendo da App.Config
            string[] mercati = Workbook.AppSettings("Mercati").Split('|');

            foreach (string msd in mercati)
            {
                Excel.Worksheet ws;
                try
                {
                    ws = Workbook.Sheets[msd];
                }
                catch
                {
                    ws = (Excel.Worksheet)Workbook.Sheets.Add(Workbook.Log);
                    ws.Name = msd;
                    ws.Select();
                    ws.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                    Workbook.Application.Windows[1].DisplayGridlines = false;
                }
            }

            return base.Struttura();
        }
        protected override void StrutturaFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.MSDSheets)
            {
                SheetExport se = new SheetExport(ws);
                se.LoadStructure();
            }

            foreach (Excel.Worksheet ws in Workbook.CategorySheets)
            {
                Sheet s = new Sheet(ws);
                s.LoadStructure();    
            }
        }
        protected override void StrutturaRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.LoadStructure();
        }

        private void AggiornaColoriVariazioni()
        {
            string mercatoPrec = Simboli.GetMercatoPrec();
            
            if (mercatoPrec != null)
            {
                DefinedNames defNamesMercatoPrec = new DefinedNames(mercatoPrec);
                DefinedNames definedNames = new DefinedNames();

                DataTable categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA];
                DataView categoria = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
                DataView categoriaEntitaView = new DataView(DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA]);
                DataView informazioni = new DataView(DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE]);

                categoriaEntitaView.RowFilter = "Gerarchia = '' OR Gerarchia IS NULL";
                string desCategoria = "";

                foreach (DataRowView entita in categoriaEntitaView)
                {
                    categoria.RowFilter = "SiglaCategoria = '" + entita["SiglaCategoria"] + "' AND Operativa = '1'";

                    if (!desCategoria.Equals(categoria[0]["DesCategoria"])) 
                    {
                        desCategoria = categoria[0]["DesCategoria"].ToString();
                        definedNames = new DefinedNames(desCategoria);
                    }
                    SplashScreen.UpdateStatus("Aggiorno i colori delle variazioni di " + entita["DesEntita"]);

                    List<DataRow> entitaRif =
                        (from r in categoriaEntita.AsEnumerable()
                         where r["Gerarchia"].Equals(entita["SiglaEntita"])
                         select r).ToList();

                    bool hasEntitaRif = entitaRif.Count > 0;
                    int numEntita = Math.Max(entitaRif.Count, 1);

                    for (int i = 0; i < numEntita; i++)
                    {
                        informazioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND Visibile = '1' " + (hasEntitaRif ? "AND SiglaEntitaRif = '" + entitaRif[i]["SiglaEntita"] + "'" : "");

                        for (int j = 0; j < informazioni.Count; j++)
                        {
                            Range rngMercatoAttuale = definedNames.Get(hasEntitaRif ? entitaRif[i]["SiglaEntita"] : entita["SiglaEntita"], informazioni[j]["SiglaInformazione"]).Extend(colOffset: Date.GetOreGiorno(DataBase.DataAttiva));

                            Range rngMercatoPrec = new Range(defNamesMercatoPrec.GetRowByName(entita["SiglaEntita"], "UM", "T") + 2, defNamesMercatoPrec.GetColFromName("RIF" + (i + 1), "PROGRAMMAQ" + (j + 1)), rowOffset: Date.GetOreGiorno(DataBase.DataAttiva));

                            for (int k = 0; k < rngMercatoAttuale.Columns.Count; k++)
                            {
                                decimal valMercatoAttuale = (decimal)(Workbook.Sheets[desCategoria].Range[rngMercatoAttuale.Columns[k].ToString()].Value ?? 0);
                                decimal valMercatoPrec = (decimal)(Workbook.Sheets[mercatoPrec].Range[rngMercatoPrec.Rows[k].ToString()].Value ?? 0);

                                if (valMercatoPrec > valMercatoAttuale)
                                {
                                    Style.RangeStyle(Workbook.Sheets[desCategoria].Range[rngMercatoAttuale.Columns[k].ToString()], backColor: 38);
                                    Style.RangeStyle(Workbook.Sheets[Simboli.Mercato].Range[rngMercatoPrec.Rows[k].ToString()], backColor: 38);
                                }
                                else if (valMercatoPrec < valMercatoAttuale)
                                {
                                    Style.RangeStyle(Workbook.Sheets[desCategoria].Range[rngMercatoAttuale.Columns[k].ToString()], backColor: 4);
                                    Style.RangeStyle(Workbook.Sheets[Simboli.Mercato].Range[rngMercatoPrec.Rows[k].ToString()], backColor: 4);
                                }
                                else
                                {
                                    Style.RangeStyle(Workbook.Sheets[desCategoria].Range[rngMercatoAttuale.Columns[k].ToString()], backColor: 2);
                                    Style.RangeStyle(Workbook.Sheets[Simboli.Mercato].Range[rngMercatoPrec.Rows[k].ToString()], backColor: 2);
                                }
                            }
                        }
                    }
                }
            }
        }


        public override bool Dati()
        {
            return base.Dati();
        }
        protected override void DatiFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.MSDSheets)
            {
                SheetExport se = new SheetExport(ws);
                se.UpdateData();
            }

            foreach (Excel.Worksheet ws in Workbook.CategorySheets)
            {
                Sheet s = new Sheet(ws);
                s.UpdateData();
            }
        }
    }
}
