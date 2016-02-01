using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Cambio i label e nascondo la riga 6.
    /// </summary>
    class Riepilogo : Base.Riepilogo
    {
        public Riepilogo()
            : base()
        {

        }

        public Riepilogo(Excel.Worksheet ws)
            : base(ws)
        {

        }

        public override void InitLabels()
        {
            base.InitLabels();

            //nascondi quelli non utilizzati
            _ws.Shapes.Item("lbImpianti").Visible = Office.MsoTriState.msoFalse;
            _ws.Shapes.Item("lbElsag").Visible = Office.MsoTriState.msoFalse;

            //sposto i due label sotto
            _ws.Shapes.Item("lbModifica").Top = _ws.Shapes.Item("lbImpianti").Top;
            _ws.Shapes.Item("lbTest").Top = _ws.Shapes.Item("lbElsag").Top;

            //ridimensiono lo sfondo
            _ws.Shapes.Item("sfondo").LockAspectRatio = Office.MsoTriState.msoFalse;
            _ws.Shapes.Item("sfondo").Height = (float)(12.5 * _ws.Rows[5].Height);
            _ws.Shapes.Item("sfondo").LockAspectRatio = Office.MsoTriState.msoTrue;
        }
        public override void UpdateData()
        {
            _ws.Shapes.Item("lbDataInizio").TextFrame.Characters().Text = Workbook.DataAttiva.ToString("ddd d MMM yyyy");
            _ws.Shapes.Item("lbDataFine").TextFrame.Characters().Text = Workbook.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("ddd d MMM yyyy");
        }

        public override void LoadStructure()
        {
            _colonnaInizio = _struttura.colRecap;
            _rigaAttiva = _struttura.rowRecap + 1;

            InitLabels();
            base.Clear();

            //if (Struct.visualizzaRiepilogo)
            //{
                _categorie.RowFilter = "Operativa = 1 AND IdApplicazione = " + Workbook.IdApplicazione;
                _azioni.RowFilter = "Visibile = 1 AND Operativa = 1 AND IdApplicazione = " + Workbook.IdApplicazione;
                _entita.RowFilter = "IdApplicazione = " + Workbook.IdApplicazione;

                CreaNomiCelle();
                //InitBarraTitolo();
                //_rigaAttiva += 2;
                FormattaRiepilogo();
                //InitBarraEntita();
                //AbilitaAzioni();
                //CaricaDatiRiepilogo();

                //Se sono in multiscreen lascio il riepilogo alla fine, altrimenti lo riporto all'inizio
                if (Screen.AllScreens.Length == 1)
                {
                    _ws.Application.ActiveWindow.SmallScroll(Type.Missing, Type.Missing, _struttura.colRecap - _struttura.colBlock - 1);
                }
                //Workbook.ScreenUpdating = false;
            //}

        }

        protected override void CreaNomiCelle()
        {
            //inserisco tutte le righe
            _definedNames.AddName(_rigaAttiva++, "TOTALE");
            _definedNames.AddName(_rigaAttiva++, "ENTITA");
            CicloGiorni((oreGiorno, suffissioData, giorno) => 
            {
                _definedNames.AddName(_rigaAttiva++, suffissioData);
            });


            //inserisco tutte le colonne
            _definedNames.AddCol(_colonnaInizio++, "GIORNI");
            foreach (DataRowView categoria in _categorie)
            {
                _entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                foreach (DataRowView entita in _entita)
                    _definedNames.AddCol(_colonnaInizio++, entita["SiglaEntita"]);
            }
            _definedNames.AddCol(_colonnaInizio++, "TOTALE");

            _definedNames.DumpToDataSet();
        }

        protected void FormattaRiepilogo()
        {
            //Titolo in alto
            Range rngTitolo = new Range(_definedNames.GetRowByName("TOTALE"), _definedNames.GetColFromName("GIORNI") + 1, 1, _definedNames.GetColOffsetRiepilogo() - 1);
            Style.RangeStyle(_ws.Range[rngTitolo.ToString()], style: "Barra titolo riepilogo", merge: true, fontSize: 10);
            _ws.Range[rngTitolo.ToString()].Value = "TOTALI";
            _ws.Range[rngTitolo.ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

            //barra delle date
            Range rngBarraDate = new Range(_definedNames.GetRowByName(Date.SuffissoDATA1), _definedNames.GetColFromName("GIORNI"), Struct.intervalloGiorni + 1);
            Style.RangeStyle(_ws.Range[rngBarraDate.ToString()], style: "Lista entita riepilogo", numberFormat: "dd/MM/yyyy", borders: "[insideh:thin]", bold: false);
            _ws.Range[rngBarraDate.ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

            //compilo i giorni
            int i = 0;
            CicloGiorni((oreGiorno, suffissoData, giorno) =>
            {
                _ws.Range[rngBarraDate.Rows[i++].ToString()].Value = giorno;
            });
            _ws.Range[rngBarraDate.ToString()].EntireColumn.AutoFit();

            //area dati disabilitata
            Range rngDati = new Range(_definedNames.GetRowByName(Date.SuffissoDATA1), _definedNames.GetColFromName("GIORNI") + 1, Struct.intervalloGiorni + 1, _definedNames.GetColOffsetRiepilogo() - 1);
            Style.RangeStyle(_ws.Range[rngDati.ToString()], style: "Area dati riepilogo", bold: false);                      //di default, le celle sono "disabilitate"
            _ws.Range[rngDati.ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

            //colonna del totale
            Range rngTotale = _definedNames.Get("ENTITA", "TOTALE").Extend(Struct.intervalloGiorni + 2);
            //titolo
            _ws.Range[rngTotale.Rows[0].ToString()].Value = "TOTALE";
            Style.RangeStyle(_ws.Range[rngTotale.ToString()], style: "Barra titolo riepilogo", fontSize: 9, borders: "[insideh:thin]", bold: true);
            //bordi totale
            _ws.Range[rngTotale.Rows[0].ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            _ws.Range[rngTotale.ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

            foreach (DataRowView categoria in _categorie)
            {
                _entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                foreach (DataRowView entita in _entita)
                {
                    Range rngEntita = _definedNames.Get("ENTITA", entita["SiglaEntita"]).Extend(Struct.intervalloGiorni + 2);
                    _ws.Range[rngEntita.ToString()].Interior.Pattern =  Excel.XlPattern.xlPatternNone;                      //riabilito celle
                    Style.RangeStyle(_ws.Range[rngEntita.Rows[0].ToString()], style: "Barra titolo riepilogo", fontSize: 9);
                    _ws.Range[rngEntita.Rows[0].ToString()].Value = entita["DesEntitaBreve"];
                }
            }

            Range rngAll = new Range(_definedNames.GetFirstRow(), _definedNames.GetFirstCol() + 1, _definedNames.GetRowOffset(), _definedNames.GetColOffsetRiepilogo());

            _ws.Range[rngAll.ToString()].ColumnWidth = 15;
        }

        protected void FormattaAllDati()
        {
            Range rngAll = new Range(_definedNames.GetFirstRow(), _definedNames.GetFirstCol() + 1, _definedNames.GetRowOffset(), _definedNames.GetColOffsetRiepilogo() - 1);
            Range rngData = new Range(_definedNames.GetFirstRow() + 2, _definedNames.GetFirstCol(), _definedNames.GetRowOffset() - 2, _definedNames.GetColOffsetRiepilogo());

            _ws.Range[rngData.ToString()].Style = "Area dati riepilogo";
            _ws.Range[rngData.Columns[0].ToString()].Style = "Lista entita riepilogo";
            _ws.Range[rngData.Columns[0].ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

            Excel.Range xlrng = _ws.Range[rngAll.Rows[1, rngAll.Rows.Count - 1].ToString()];
            //trovo tutte le aree unite e creo il blocco col bordo grosso
            int i = 0;
            int colspan = 0;
            while (i < xlrng.Columns.Count)
            {
                colspan = xlrng.Cells[1, i + 1].MergeArea().Columns.Count;
                _ws.Range[rngAll.Columns[i, i + colspan - 1].ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
                _ws.Range[rngAll.Columns[i, i + colspan - 1].ToString()].Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
                i += colspan;
            }


            _ws.Range[rngAll.ToString()].EntireColumn.AutoFit();
            if (rngAll.ColOffset > 1)
            {
                //calcolo la massima dimensione delle colonne e la riapplico a tutto il riepilogo
                double maxWidth = double.MinValue;
                foreach (Range col in rngAll.Columns)
                    maxWidth = Math.Max(_ws.Range[col.ToString()].ColumnWidth, maxWidth);

                foreach (Range col in rngAll.Columns)
                    _ws.Range[col.ToString()].ColumnWidth = maxWidth;
            }
        }

        //protected void InitBarraTitolo()
        //{
        //    Range rngTitleBar = new Range(_definedNames.GetFirstRow(), _definedNames.GetFirstCol() + 1, 2, _categorie.Count);
        //    Range rngAll = new Range(_definedNames.GetFirstRow() + 1, _definedNames.GetFirstCol() + 1, _definedNames.GetRowOffset() - 1, _definedNames.GetColOffsetRiepilogo() - 1);
        //    //Range rngData = rngTitleBar.Cells[0, 0];
        //    //Range rngEntita = rngTitleBar.Cells[1, 0];

        //    Style.RangeStyle(_ws.Range[rngTitleBar.Rows[0].ToString()], style: "Barra titolo riepilogo", merge: true);
        //    _ws.Range[rngTitleBar.Rows[0].ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

        //    foreach (DataRowView categoria in _categorie)
        //    {
        //        _entita.RowFilter = "SiglaCategoria = '" + categoria["SiglaCategoria"] + "' AND IdApplicazione = " + Workbook.IdApplicazione;
        //        foreach (DataRowView entita in _entita)
        //        {
                    
                    
                    
        //            //entita["DesEntita"]
        //        }
        //            //_definedNames.AddCol(_colonnaInizio++, entita["SiglaEntita"]);
        //    }
            
            
            
            
        //    string azionePadre = "";
            
            
            
        //    //CicloGiorni((oreGiorno, suffissoData, giorno) =>
        //    //{
        //    //    rngTitleBar.StartColumn = rngAzioni.StartColumn;
        //    //    _ws.Range[rngTitleBar.ToString()].Style = "Barra titolo riepilogo";
        //    //    _ws.Range[rngTitleBar.ToString()].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);

        //    //    foreach (DataRowView azione in _azioni)
        //    //    {
        //    //        if (!azione["Gerarchia"].Equals(azionePadre))
        //    //        {
        //    //            rngEntita.ColOffset = rngAzioni.StartColumn - rngEntita.StartColumn;
        //    //            Style.RangeStyle(_ws.Range[rngEntita.ToString()], merge: true, fontSize: 9);
        //    //            _ws.Range[rngEntita.ToString()].Value = azionePadre;
        //    //            azionePadre = azione["Gerarchia"].ToString();
        //    //            rngEntita.StartColumn = rngAzioni.StartColumn;
        //    //        }
        //    //        _ws.Range[rngAzioni.ToString()].Value = azione["DesAzioneBreve"];
        //    //        Style.RangeStyle(_ws.Range[rngAzioni.ToString()], fontSize: 7);
        //    //        rngAzioni.StartColumn++;
        //    //    }
        //    //    rngEntita.ColOffset = rngAzioni.StartColumn - rngEntita.StartColumn;
        //    //    Style.RangeStyle(_ws.Range[rngEntita.ToString()], merge: true, fontSize: 9);
        //    //    _ws.Range[rngEntita.ToString()].Value = azionePadre;
        //    //    azionePadre = "";
        //    //    rngEntita.StartColumn = rngAzioni.StartColumn;

        //    //    rngData.ColOffset = rngAzioni.StartColumn - rngData.StartColumn;
        //    //    Style.RangeStyle(_ws.Range[rngData.ToString()], merge: true, fontSize: 10, numberFormat: "ddd d mmm yyyy");
        //    //    _ws.Range[rngData.ToString()].Value = giorno;
        //    //    rngData.StartColumn = rngAzioni.StartColumn;
        //    //});

        //    UpdateDayColor();
        //}
    }
}
