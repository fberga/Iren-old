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
        ImageList _imgList;

        public ConfiguratoreTasto(Button btn, ImageList imgList)
        {
            InitializeComponent();

            _btn = btn;
            _imgList = imgList;

            imgButton.ImageLocation = _btn.ImageKey;
            txtName.Text = _btn.Name;
            txtLabel.Text = _btn.Text;
            
            Dictionary<string, object> metaData = _btn.Tag as Dictionary<string, object>;

            object desc;
            object screenTip;
            object toggleBtn;
            object size;

            metaData.TryGetValue(ConfiguratoreRibbon.DESC_FIELD_NAME, out desc);
            metaData.TryGetValue(ConfiguratoreRibbon.SCREEN_TIP_FIELD_NAME, out screenTip);
            metaData.TryGetValue(ConfiguratoreRibbon.TOGGLE_BUTTON_FIELD_NAME, out toggleBtn);
            metaData.TryGetValue(ConfiguratoreRibbon.DIMENSION_FIELD_NAME, out size);

            txtDesc.Text = desc as string ?? "";
            txtScreenTip.Text = screenTip as string ?? "";
            chkToggleButton.Checked = (bool)(toggleBtn ?? false);            
            if ((int)(size ?? 1) == 0)
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
            _btn.Name = txtName.Text;
            _btn.Text = txtLabel.Text;

            Dictionary<string, object> metaData = _btn.Tag as Dictionary<string, object>;
            metaData[ConfiguratoreRibbon.DESC_FIELD_NAME] = txtDesc.Text;
            metaData[ConfiguratoreRibbon.SCREEN_TIP_FIELD_NAME] = txtScreenTip.Text;
            metaData[ConfiguratoreRibbon.TOGGLE_BUTTON_FIELD_NAME] = chkToggleButton.Checked;
            metaData[ConfiguratoreRibbon.DIMENSION_FIELD_NAME] = radioDimSmall.Checked ? 0 : 1;

            Close();
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
