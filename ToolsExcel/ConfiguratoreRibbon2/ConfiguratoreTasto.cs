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
        RibbonButton _btn;
        ImageList _imgList;

        public ConfiguratoreTasto(RibbonButton btn, ImageList imgList)
        {
            InitializeComponent();

            _btn = btn;
            _imgList = imgList;

            imgButton.ImageLocation = _btn.ImageKey;
            txtName.Text = _btn.Label;
            txtLabel.Text = _btn.Nome;

            txtDesc.Text = _btn.Descrizione;
            txtScreenTip.Text = _btn.ScreenTip;
            chkToggleButton.Checked = _btn.ToggleButton;            
            if (_btn.Dimensione == 0)
                radioDimSmall.Checked = true;
            else
                radioDimLarge.Checked = true;
        }

        private void ChangeBtnImage(object sender, EventArgs e)
        {
            using (SelettoreImmagini chooseImageDialog = new SelettoreImmagini(_imgList))
            {
                if (chooseImageDialog.ShowDialog() == DialogResult.OK)
                    imgButton.ImageLocation = chooseImageDialog.FileName;
            }
        }

        private void Applica_Click(object sender, EventArgs e)
        {
            _btn.ImageKey = imgButton.ImageLocation;
            _btn.Nome = txtName.Text;
            _btn.Label = txtLabel.Text;

            _btn.Descrizione = txtDesc.Text;
            _btn.ScreenTip = txtScreenTip.Text;
            _btn.ToggleButton = chkToggleButton.Checked;
            _btn.Dimensione = radioDimSmall.Checked ? 0 : 1;

            Close();
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
