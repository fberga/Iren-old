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
    class SheetExport : Base.ASheet
    {
        protected Excel.Worksheet _ws;
        protected DefinedNames _definedNames;

        public SheetExport(Excel.Worksheet ws)
        {
            _ws = ws;

            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            categorie.RowFilter = "DesCategoria = '" + ws.Name + "'";

            AggiornaParametriSheet();

            _definedNames = new DefinedNames(_ws.Name);
        }

        protected void AggiornaParametriSheet()
        {
            DataView paramApplicazione = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE].DefaultView;

            _struttura = new Struct();

            _struttura.rigaBlock = (int)paramApplicazione[0]["RowBlocco"];
            _struttura.rigaGoto = (int)paramApplicazione[0]["RowGoto"];

        }

        public override void LoadStructure()
        {
            Clear();
        }

        private void Clear()
        {
            if (_ws.ChartObjects().Count > 0)
                _ws.ChartObjects().Delete();

            _ws.UsedRange.EntireColumn.Delete();
            _ws.UsedRange.FormatConditions.Delete();
            _ws.UsedRange.EntireRow.Hidden = false;
            _ws.UsedRange.Font.Size = 10;
            _ws.UsedRange.NumberFormat = "General";
            _ws.UsedRange.Font.Name = "Verdana";
            _ws.UsedRange.RowHeight = Struct.cell.height.normal;

            _ws.Columns.ColumnWidth = Struct.cell.width.dato;

            _ws.Rows["1:" + (_struttura.rigaBlock - 1)].RowHeight = Struct.cell.height.empty;

            _ws.Rows[_struttura.rigaGoto].RowHeight = Struct.cell.height.normal;

            _ws.Columns[1].ColumnWidth = Struct.cell.width.empty;

            ((Excel._Worksheet)_ws).Activate();
            _ws.Application.ActiveWindow.FreezePanes = false;
            _ws.Cells[_struttura.rigaBlock, 1].Select();
            _ws.Application.ActiveWindow.ScrollColumn = 1;
            _ws.Application.ActiveWindow.ScrollRow = 1;
            _ws.Application.ActiveWindow.FreezePanes = true;
            Workbook.Main.Select();
            _ws.Application.ScreenUpdating = false;
        }

        public override void UpdateData(bool all = true)
        {
            throw new NotImplementedException();
        }

        public override void AggiornaDateTitoli()
        {
            throw new NotImplementedException();
        }

        public override void AggiornaGrafici()
        {
            throw new NotImplementedException();
        }

        protected override void InsertPersonalizzazioni(object siglaEntita)
        {
            throw new NotImplementedException();
        }

        public override void CaricaInformazioni(bool all)
        {
            throw new NotImplementedException();
        }
    }
}
