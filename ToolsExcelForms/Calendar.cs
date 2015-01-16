using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Iren.FrontOffice.Forms
{
    public partial class frmCALENDAR : Form
    {
        public string date = "";

        public frmCALENDAR()
        {
            InitializeComponent();
            Application.EnableVisualStyles();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            //_t.SetParametroApplicazione(new Dictionary<string, object>() { { FOConst.PARAMETERS.DATA_SELEZIONATA, calObj.SelectionStart.ToString("yyyyMMdd") } });
            date = calObj.SelectionStart.ToString("yyyyMMdd");
            this.Close();
        }

        private void btnANNULLA_Click(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}
