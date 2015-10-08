using Iren.ToolsExcel.Utility;
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
    public partial class ControlliEsistenti : Form
    {
        

        DataTable _dtCtrl;
        ImageList _imgList;

        public ControlliEsistenti(ImageList imgSmall, ImageList imgNormal)
        {
            InitializeComponent();
            imgSmall.Images.Add("emptyImage", new Bitmap(1, 1));
            treeViewControlli.ImageList = imgSmall;
            _imgList = imgNormal;
        }

        public ControlliEsistenti(ImageList imgSmall, ImageList imgNormal, params int[] controlType)
            : this(imgSmall, imgNormal)
        {
            _dtCtrl = DataBase.Select(SP.CONTROLLO);

            var typesDesc = _dtCtrl.AsEnumerable()
                .Where(r => controlType.Contains((int)r["IdTipologiaControllo"]))
                .Select(r => new { ID = r["IdTipologiaControllo"], Desc = r["DesTipologiaControllo"] })
                .Distinct();

            foreach (var type in typesDesc)
            {
                TreeNode typeRoot = new TreeNode(type.Desc.ToString());
                typeRoot.ImageKey = "emptyImage";
                typeRoot.SelectedImageKey = "emptyImage";

                var controls = _dtCtrl.AsEnumerable()
                    .Where(r => r["IdTipologiaControllo"].Equals(type.ID));

                foreach (var ctrl in controls)
                {
                    TreeNode c = new TreeNode(ctrl["Label"].ToString());

                    c.Tag = ctrl["IdControllo"];

                    if (!ctrl["Immagine"].Equals(""))
                    {
                        c.ImageKey = ctrl["Immagine"].ToString();
                        c.SelectedImageKey = ctrl["Immagine"].ToString();
                    }
                    else
                    {
                        c.ImageKey = "emptyImage";
                        c.SelectedImageKey = "emptyImage";
                    }

                    typeRoot.Nodes.Add(c);
                }

                treeViewControlli.Nodes.Add(typeRoot);
            }
            treeViewControlli.ExpandAll();
        }

        private void AfterSelectNode(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Tag != null && e.Node.Tag.GetType() == typeof(int))
            {
                int id = (int)e.Node.Tag;

                var selectedCtrl = _dtCtrl.AsEnumerable()
                    .Where(c => c["IdControllo"].Equals(id))
                    .FirstOrDefault();

                imgButton.Image = _imgList.Images[selectedCtrl["Immagine"].ToString()];
                txtDesc.Text = selectedCtrl["Descrizione"].ToString();
                txtScreenTip.Text = selectedCtrl["ScreenTip"].ToString();
                
                //carico informazioni su gruppi, applicazioni e funzioni utilizzate
                DataTable ribbons = DataBase.Select(SP.APPLICAZIONE_UTENTE_RIBBON);

                DataView gruppi = new DataView(ribbons);
                DataView applicazioni = new DataView(ribbons);

                gruppi.RowFilter = "IdControllo = " + id;
                listBoxGruppi.DisplayMember = "LabelGruppo";
                listBoxGruppi.ValueMember = "IdGruppo";

                applicazioni.RowFilter = "IdGruppo = -1";
                listBoxApplicazioni.DisplayMember = "DesApplicazione";


                listBoxGruppi.DataSource = gruppi;
                listBoxApplicazioni.DataSource = applicazioni;

                if(listBoxGruppi.Items.Count > 0)
                    SelectedGroupChanged(listBoxGruppi, new EventArgs());
            }
        }

        private void SelectedGroupChanged(object sender, EventArgs e)
        {
            if (listBoxApplicazioni.DataSource != null)
            {
                DataView dvG = listBoxGruppi.DataSource as DataView;
                DataView dvA = listBoxApplicazioni.DataSource as DataView;
                
                dvA.RowFilter = dvG.RowFilter + " AND IdGruppo = " + listBoxGruppi.SelectedValue;
            }
        }
    }
}
