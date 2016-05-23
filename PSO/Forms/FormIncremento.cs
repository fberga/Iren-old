using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Forms
{
    public partial class FormIncremento : Form
    {
        private object[] _origVal;
        private Range _origRng;
        private Excel.Range _toModify;
        private DefinedNames _definedNames;

        public FormIncremento(Excel.Worksheet ws, Excel.Range rng)
        {
            InitializeComponent();
            this.Text = Simboli.NomeApplicazione + " - Incremento";

            _origRng = new Range(rng.Address);
            int marketOffset = Simboli.GetMarketOffset(DateTime.Now.Hour);

            _definedNames = new DefinedNames(ws.Name);

            Range rowRange = new Range(rng.Row, _definedNames.GetColFromDate(Date.SuffissoDATA1), 1, Date.GetOreGiorno(Workbook.DataAttiva));


            if (_origRng.StartColumn < rowRange.StartColumn + marketOffset)
            {
                if (MessageBox.Show("Il range selezionato contiene celle non modificabili. Per continuare è necessario modificare la selezione. Premendo Ok verrà selezionato il range modificabile più vicino a quello selezionato. Annullando l'operazione si potrà procedere alla modifica manuale.", Simboli.NomeApplicazione + " - ATTENZIONE!!!", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
                {
                    this.Close();
                    return;
                }

                _origRng.ColOffset -= rowRange.StartColumn + marketOffset - _origRng.StartColumn;
                _origRng.StartColumn = rowRange.StartColumn + marketOffset;

            }

            txtRangeSelezionato.Text = _origRng.ToString();
            _toModify = rng;
        }

        private void chkTuttaRiga_CheckedChanged(object sender, EventArgs e)
        {
            if (chkTuttaRiga.Checked)
            {

            }
            else
            {

            }
        }
    }
}
