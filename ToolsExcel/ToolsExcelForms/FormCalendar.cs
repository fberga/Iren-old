using Iren.ToolsExcel.Base;
using System;
using System.Windows.Forms;

namespace Iren.ToolsExcel.Forms
{
    public partial class FormCalendar : Form
    {
        public FormCalendar()
        {
            InitializeComponent();
            Application.EnableVisualStyles();
            calObj.SetDate(Utility.DataBase.DataAttiva);
            this.Text = Simboli.nomeApplicazione + " - Calendar";
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnANNULLA_Click(object sender, EventArgs e)
        {
            calObj.SetDate(Utility.DataBase.DataAttiva);
            this.Close();
        }

        /// <summary>
        /// Override del metodo ShowDialog di Windows Forms. Restituisce l'oggetto data selezionato se l'utente ha cambiato la selezione, null altrimenti.
        /// </summary>
        /// <returns>Restituisce l'oggetto data selezionato se l'utente ha cambiato la selezione, null altrimenti.</returns>
        public new object ShowDialog()
        {
            base.ShowDialog();
            if (calObj.SelectionStart == Utility.DataBase.DataAttiva)
                return null;

            return calObj.SelectionStart;
        }

    }
}
