using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System.Data;
using System.Globalization;

namespace ProvaRibbon
{
    public partial class AutoRibbon
    {
        public void InitializeComponent2()
        {
            //EventInfo ei = btnCalendar.GetType().GetEvent("Click");
            //MethodInfo hi = GetType().GetMethod("btnCalendar_Click", BindingFlags.Instance | BindingFlags.NonPublic);
            //Delegate d = Delegate.CreateDelegate(ei.EventHandlerType, null, hi);
            //ei.AddEventHandler(btnCalendar, d);


            //this.groupChiudi = this.Factory.CreateRibbonGroup();
            //this.btnEsportaXML = this.Factory.CreateRibbonButton();
            //this.btnValidazioneTL = this.Factory.CreateRibbonToggleButton();
            //this.labelMSD = this.Factory.CreateRibbonLabel();
            //this.cmbMSD = this.Factory.CreateRibbonComboBox();

            //this.btnChiudi.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            //this.btnChiudi.Image = ToolsExcel.Base.Properties.Resources.chiudi;
            //this.btnChiudi.Label = "Chiudi";
            //this.btnChiudi.Name = "btnChiudi";
            //this.btnChiudi.ShowImage = true;

            DataBase.InitNewDB("Dev");

            if (DataBase.OpenConnection())
            {
                //TODO salvare anche questo negli XML
                DataTable dt = DataBase.Select("RIBBON.spGruppoControllo", "@IdApplicazione=1;@IdUtente=62");

                Microsoft.Office.Tools.Ribbon.RibbonGroup grp = this.Factory.CreateRibbonGroup();
                string nomeGruppo = "";

                foreach (DataRow r in dt.Rows)
                {
                    if (!r["NomeGruppo"].Equals(nomeGruppo))
                    {
                        nomeGruppo = r["NomeGruppo"].ToString();
                        grp = this.Factory.CreateRibbonGroup();
                        grp.Name = nomeGruppo;
                        grp.Label = r["LabelGruppo"].ToString();

                        this.FrontOffice.Groups.Add(grp);
                    }

                    if(typeof(RibbonButton).FullName.Equals(r["SiglaTipologiaControllo"])) 
                    {
                        RibbonButton newBtn = this.Factory.CreateRibbonButton();

                        newBtn.ControlSize = (Microsoft.Office.Core.RibbonControlSize)r["ControlSize"];
                        newBtn.Image = (System.Drawing.Image)Iren.ToolsExcel.Base.Properties.Resources.ResourceManager.GetObject(r["Immagine"].ToString());
                        newBtn.Label = r["Label"].ToString();
                        newBtn.ShowImage = true;
                        
                        grp.Items.Add(newBtn);
                    }
                    else if (typeof(RibbonToggleButton).FullName.Equals(r["SiglaTipologiaControllo"])) 
                    {
                        RibbonToggleButton newTglBtn = this.Factory.CreateRibbonToggleButton();

                        newTglBtn.ControlSize = (Microsoft.Office.Core.RibbonControlSize)r["ControlSize"];
                        newTglBtn.Image = (System.Drawing.Image)Iren.ToolsExcel.Base.Properties.Resources.ResourceManager.GetObject(r["Immagine"].ToString());
                        newTglBtn.Label = r["Label"].ToString();
                        newTglBtn.ShowImage = true;

                        grp.Items.Add(newTglBtn);
                    }
                    else if (typeof(RibbonComboBox).FullName.Equals(r["SiglaTipologiaControllo"])) 
                    {
                        RibbonLabel lb = this.Factory.CreateRibbonLabel();
                        lb.Label = r["Label"].ToString();
                        RibbonComboBox cmb = this.Factory.CreateRibbonComboBox();
                        cmb.ShowLabel = false;
                        cmb.Text = null;

                        cmb.ItemsLoading += cmb_ItemsLoading;

                        grp.Items.Add(lb);
                        grp.Items.Add(cmb);
                    }
                }

            }



        }

        void cmb_ItemsLoading(object sender, RibbonControlEventArgs e)
        {
            RibbonComboBox ccc = sender as RibbonComboBox;

            if (ccc.Items.Count == 0)
            {
                RibbonDropDownItem i = Factory.CreateRibbonDropDownItem();
                i.Label = "Primavera";
                ccc.Items.Add(i);
                i = Factory.CreateRibbonDropDownItem();
                i.Label = "Estate";
                ccc.Items.Add(i);
                i = Factory.CreateRibbonDropDownItem();
                i.Label = "Autunno";
                ccc.Items.Add(i);
                i = Factory.CreateRibbonDropDownItem();
                i.Label = "Inverno";
                ccc.Items.Add(i);

                ccc.Text = i.Label;
            }
        }


        private void AutoRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }
    }
}
