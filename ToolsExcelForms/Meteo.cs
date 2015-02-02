using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Iren.FrontOffice.Core;
using System.Globalization;

namespace Iren.FrontOffice.Forms
{
    public partial class frmMETEO : Form
    {
        DataBase _db;
        DataView _entita;
        DataView _entitaProprieta;
        DateTime _dataRif;


        public frmMETEO(DataView entita, DataView entitaProprieta, object dataRif, DataBase db)
        {
            InitializeComponent();

            _db = db;
            _entita = entita;
            _entitaProprieta = entitaProprieta;
            _dataRif = (DateTime)dataRif;

            labelData.Text = "Data Riferimento: " + _dataRif.ToString("dd/MM/yyyy");
        }

        private void frmMETEO_Load(object sender, EventArgs e)
        {
            _entitaProprieta.RowFilter = "SiglaProprieta = 'PROGR_IMPIANTO_TEMP_FONTE_ATTIVA'";


            comboUP.DataSource = _entita;
            comboUP.DisplayMember = "DesEntita";

        }

        private void comboUP_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_db.StatoDB()[DataBase.NomiDB.SQLSERVER] == ConnectionState.Open)
            {
                comboARPA.Items.Clear();
                comboEPSON.Items.Clear();
                comboNIMBUS.Items.Clear();

                DataView fonti = _db.Select("spCheckFonteMeteo", "@SiglaEntita=" + ((DataRowView)comboUP.SelectedItem)["SiglaEntita"] + ";@Data=" + _dataRif.ToString("yyyyMMdd")).DefaultView;

                foreach (DataRowView fonte in fonti)
                {
                    DateTime dataEmissione = DateTime.ParseExact(fonte["DataEmissione"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                    switch (fonte["CodiceFonte"].ToString())
                    {
                        case "ARPA":
                            comboARPA.Items.Add(dataEmissione);
                            break;
                        case "EPSON":
                            comboEPSON.Items.Add(dataEmissione);
                            break;
                        case "NIMBUS":
                            comboNIMBUS.Items.Add(dataEmissione);
                            break;
                    }
                }

                foreach (ComboBox cmb in groupDati.Controls.OfType<ComboBox>())
                {
                    string name = cmb.Name.Replace("combo", "radio");
                    RadioButton rd = (RadioButton)groupDati.Controls[name];
                    rd.Checked = false;
                    
                    if (cmb.Items.Count > 0) 
                    {
                        cmb.SelectedIndex = 0;
                        cmb.Visible = true;
                        rd.Visible = true;
                    }  
                    else
                    {
                        cmb.Visible = false;
                        rd.Visible = false;
                    }
                }
                if (groupDati.Controls.OfType<ComboBox>().Where(cmb => cmb.Visible).ToArray().Length == 0)
                {
                    labelDataEmissione.Visible = false;
                }
                else
                {
                    labelDataEmissione.Visible = true;
                }
            }
        }

        private void comboDataEmissione_DataSourceChanged(object sender, EventArgs e)
        {
            ComboBox cmb = (ComboBox)sender;
            cmb.Visible = cmb.Items.Count != 0;
            
            string name = cmb.Name.Replace("combo", "radio");

            RadioButton rd = (RadioButton)Controls[name];
            if (rd.Checked)
                rd.Checked = false;
            
            rd.Visible = false;
        }


    }
}
