using Iren.ToolsExcel.Base;
using System;
using System.Windows.Forms;

namespace Iren.ToolsExcel.Forms
{
    public partial class FormCalendar : Form
    {
        private DateTime? _date = null;

        public DateTime? Date 
        { 
            get 
            { 
                return _date; 
            } 
        }

        public FormCalendar()
        {
            InitializeComponent();
            Application.EnableVisualStyles();
            calObj.SetDate(Utility.DataBase.DataAttiva);
            this.Text = Simboli.nomeApplicazione + " - Calendar";
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            _date = calObj.SelectionStart;
            this.Close();
        }

        private void btnANNULLA_Click(object sender, EventArgs e)
        {
            _date = null;
            this.Close();
        }


    }
}
