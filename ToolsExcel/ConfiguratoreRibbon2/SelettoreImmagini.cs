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
    public partial class SelettoreImmagini : Form
    {
        ImageList _imgs;

        public string ResourceName { get; private set; }
        public int Index { get; private set; }
        public Image Img { get; private set; }

        public SelettoreImmagini(ImageList imgs)
        {
            InitializeComponent();

            _imgs = imgs;

            imageListView.LargeImageList = _imgs;

            int i = 0;
            foreach(string img in _imgs.Images.Keys)
            {
                ListViewItem item = new ListViewItem();
                item.Text = img;
                item.ToolTipText = img;
                item.ImageIndex = i++;
                item.ImageKey = img;
                imageListView.Items.Add(item);
            }
        }

        private void SelectItemByDoubleClick(object sender, MouseEventArgs e)
        {
            if(e.Button == System.Windows.Forms.MouseButtons.Left) 
            {
                Applica_Click(null, null);
            }
        }

        private void Applica_Click(object sender, EventArgs e)
        {
            if (imageListView.SelectedItems.Count > 0)
            {
                ResourceName = imageListView.SelectedItems[0].ImageKey;
                Index = imageListView.SelectedIndices[0];
                Img = Utility.GetResurceImage(ResourceName);
                DialogResult = System.Windows.Forms.DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show(char.ToUpper('è') + " necessario selezionare un'immagine prima di proseguire...", "ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public new DialogResult ShowDialog()
        {
            if(base.ShowDialog() == DialogResult.Cancel)
                return DialogResult.Cancel;

            if (imageListView.SelectedIndices.Count == 1)
                return DialogResult.OK;

            return DialogResult.Cancel;
        }

        private void Annulla_Click(object sender, EventArgs e)
        {
            imageListView.SelectedIndices.Clear();
            
            DialogResult = System.Windows.Forms.DialogResult.Cancel;
            
            Close();
        }
    }
}
