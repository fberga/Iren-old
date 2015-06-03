using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Iren.ToolsExcel.Base;


namespace Iren.ToolsExcel.Forms
{
    public partial class FormModificaParametri : Form
    {
        DataView _parametriD;
        DataView _parametriH;
        DataView _entita;

        public FormModificaParametri()
        {
            InitializeComponent();

            _parametriD = new DataView(Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PARAMETRO_D]);
            _parametriH = new DataView(Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PARAMETRO_H]);
            _entita = new DataView(Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.CATEGORIA_ENTITA]);

            cmbEntita.ValueMember = "SiglaEntita";
            cmbEntita.DisplayMember = "DesEntita";
            cmbEntita.DataSource = _entita;

            cmbParametriD.DisplayMember = "SiglaParametro";
            cmbParametriD.DataSource = _parametriD;
            
            cmbParametriH.DisplayMember = "SiglaParametro";
            cmbParametriH.DataSource = _parametriH;
        }

        private void cmbEntita_SelectedIndexChanged(object sender, EventArgs e)
        {
            _parametriD.RowFilter = "SiglaEntita = '" + cmbEntita.SelectedValue.ToString() + "'";
            _parametriH.RowFilter = "SiglaEntita = '" + cmbEntita.SelectedValue.ToString() + "'";
        }

        private void cmbParametriD_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
