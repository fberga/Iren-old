using Iren.FrontOffice.UserConfig;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Iren.FrontOffice.Base;

namespace Iren.FrontOffice.Forms
{
    public partial class FormConfig : Form
    {
        DataTable _dt = new DataTable("usrConfig") 
        {
            Columns =
            {
                {"Key", typeof(string)},
                {"Proprietà", typeof(string)},
                {"Valore", typeof(string)},
                {"Default", typeof(string)}
            }
        };

        public FormConfig()
        {
            InitializeComponent();

            dataGridConfigurazioni.DataSource = _dt;
            
            this.Text = Simboli.nomeApplicazione + " - Config";

            int width = dataGridConfigurazioni.Width * 90 / 100;

            dataGridConfigurazioni.Columns[0].Visible = false;

            dataGridConfigurazioni.Columns[1].Width = (width / 3);
            dataGridConfigurazioni.Columns[1].ReadOnly = true;
            dataGridConfigurazioni.Columns[1].DefaultCellStyle = new DataGridViewCellStyle() 
            {
                SelectionBackColor = System.Drawing.Color.White,
                SelectionForeColor = System.Drawing.Color.Black,
                Font = new Font(dataGridConfigurazioni.Font, FontStyle.Bold)
            };
            
            dataGridConfigurazioni.Columns[2].Width = (width / 3);
            
            dataGridConfigurazioni.Columns[3].Width = (width / 3);
            dataGridConfigurazioni.Columns[3].ReadOnly = true;
            
        }

        private void FormConfig_Load(object sender, EventArgs e)
        {
            var config = (UserConfiguration)ConfigurationManager.GetSection("usrConfig");

            foreach (UserConfigElement item in config.Items)
                _dt.Rows.Add(item.Key, item.Desc, item.Value, item.Default);

        }

        private void btnApplica_Click(object sender, EventArgs e)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var section = (UserConfiguration)config.GetSection("usrConfig");

            foreach (DataRow r in _dt.GetChanges().Rows)
                section.Items[r["Key"].ToString()].Value = r["Valore"].ToString();

            config.Save(ConfigurationSaveMode.Minimal);
            ConfigurationManager.RefreshSection("usrConfig");
        }
    }
}
