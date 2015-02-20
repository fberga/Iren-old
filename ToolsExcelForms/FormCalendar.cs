using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Iren.FrontOffice.Base;

namespace Iren.FrontOffice.Forms
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
