using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ConfiguratoreRibbon2
{
    public partial class ConfiguratoreTasto : Form
    {
        Button _btn;

        public ConfiguratoreTasto(Button btn)
        {
            InitializeComponent();

            chooseImageDialog.InitialDirectory = @"D:\Repository\Iren\ToolsExcel\ToolsExcelBase\resources";
            chooseImageDialog.Filter = "PNG Files (*.png)|*.png";

            _btn = btn;

            imgButton.Image = btn.Image;

        }

        private void ChangeBtnImage(object sender, EventArgs e)
        {
            if (chooseImageDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                imgButton.ImageLocation = chooseImageDialog.FileName;
            }
        }
    }
}
