using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public partial class ConfiguratoreTasto : Form
    {
        RibbonButton _btn;
        ImageList _imgList;
        bool _applica = false;


        public ConfiguratoreTasto(RibbonButton btn, ImageList imgList)
        {
            InitializeComponent();

            _btn = btn;
            _imgList = imgList;

            imgButton.ImageLocation = _btn.ImageKey;
            //txtName.Text = _btn.Nome;
            txtLabel.Text = _btn.Label;

            txtDesc.Text = _btn.Descrizione;
            txtScreenTip.Text = _btn.ScreenTip;
            chkToggleButton.Checked = _btn.ToggleButton;            
            if (_btn.Slot == 1) 
            {
                radioDimSmall.Checked = true;
                ControlContainer ctrl = _btn.Parent as ControlContainer;
                if (ctrl.CtrlCount > 1)
                    radioDimLarge.Enabled = false;
            }
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
            //_btn.Nome = txtName.Text;
            _btn.Label = txtLabel.Text;

            _btn.Descrizione = txtDesc.Text;
            _btn.ScreenTip = txtScreenTip.Text;
            _btn.ToggleButton = chkToggleButton.Checked;
            _btn.Dimensione = radioDimSmall.Checked ? 0 : 1;

            _applica = true;
            Close();
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            Close();
        }

        public new DialogResult ShowDialog()
        {
            if (base.ShowDialog() == DialogResult.Cancel)
                return DialogResult.Cancel;

            if(_applica == true)
                return DialogResult.OK;

            return DialogResult.Cancel;
        }
    }
}
