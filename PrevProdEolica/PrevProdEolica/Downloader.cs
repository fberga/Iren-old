using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Iren.EpexDownloader
{
    public partial class Downloader : Form
    {
        private WebClient _webClient = new WebClient();
        private HtmlAgilityPack.HtmlDocument _htmlDoc = new HtmlAgilityPack.HtmlDocument();
        private string _baseURL = "http://www.terna.it";
        private string _dwnldURL = "/default/Home/SISTEMA_ELETTRICO/transparency_report/Generation/Forecast_generation_wind.aspx";
        private string _basePath = @"D:\Users\e-bergamin\Desktop";
        private DateTime _data;


        public Downloader()
        {
            InitializeComponent();

            _basePath = ConfigurationManager.AppSettings["basePath"] ?? _basePath;
            _baseURL = ConfigurationManager.AppSettings["baseURL"] ?? _baseURL;
            _dwnldURL = ConfigurationManager.AppSettings["dwnldURL"] ?? _dwnldURL;

            _data = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

            try
            {
                _htmlDoc.LoadHtml(_webClient.DownloadString(_baseURL + _dwnldURL));

                //ottengo l'array delle date visualizzate
                HtmlNodeCollection nodes = _htmlDoc.DocumentNode.SelectNodes("//div[@class='DNN_Documents']//table//tr");

                foreach (var node in nodes)
                {
                    if (node.SelectSingleNode(".//td[@class='OwnerCell']") != null
                        && node.SelectSingleNode(".//td[@class='OwnerCell']").InnerText == "Previsione Produzione Eolica"
                        && node.SelectSingleNode(".//td[@class='CategoryCell']").InnerText == _data.ToString("dd/MM/yyyy"))
                    {
                        string link = node.SelectSingleNode("//td[@class='OwnerCell']//a").Attributes["href"].Value;
                        webBrowser1.Navigate(_baseURL + link);
                        break;
                    }
                }
            }
            catch (Exception)
            {

            }
        }
    }
}
