using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Aggiungo la personalizzazione delle note.
    /// </summary>
    class Sheet : Base.Sheet
    {
        public Sheet(Excel.Worksheet ws)
            : base(ws) 
        { 
        
        }

        protected override void InsertPersonalizzazioni(object siglaEntita)
        {
            //da classe base il filtro è corretto
            DataView informazioni = Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].DefaultView;

            _ws.Columns[3].Font.Size = 9;

            int col = _definedNames.GetFirstCol();
            int row = _definedNames.GetRowByName(siglaEntita, "T");

            //metto cella con scritta totale            
            Excel.Range title = _ws.Range[Range.GetRange(row, col + 25)];
            title.Value = "TOTALE";


            Excel.Range rngPersonalizzazioni = _ws.Range[Range.GetRange(row, col + 25, _intervalloGiorniMax + 1)];

            rngPersonalizzazioni.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;

            title.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            Style.RangeStyle(title, bold: true, backColor: 8, align: Excel.XlHAlign.xlHAlignCenter);
            
            rngPersonalizzazioni.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            rngPersonalizzazioni.Columns[1].ColumnWidth = Struct.cell.width.jolly1;
            
            int i = 1;
            CicloGiorni(_dataInizio, _dataInizio.AddDays(Struct.intervalloGiorni - 1), (oreGiorno, suffissoData, giorno) => 
            {
                //_definedNames.AddName(row + i++, siglaEntita, "TOTALE", suffissoData);
                //_definedNames.SetEditable(row + i, new Range(row + i, col + 25));

                int gasDayStart = TimeZone.CurrentTimeZone.IsDaylightSavingTime(giorno) ? 7 : 6;
                int remainingHours = 24 - Date.GetOreGiorno(giorno) + gasDayStart;

                Range rng1 = new Range(row + i, _definedNames.GetColData1H1() + gasDayStart - 1, 1, 25 - gasDayStart + 1);
                Range rng2 = new Range(row + i + 1, _definedNames.GetColData1H1(), 1, remainingHours - 1);

                rngPersonalizzazioni.Cells[i + 1, 1].Formula = "=SUM("+rng1.ToString()+") + SUM("+rng2.ToString()+")";
                i++;
            });
        }
    }
}
