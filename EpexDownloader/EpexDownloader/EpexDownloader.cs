using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using HtmlAgilityPack;
using System.Globalization;
using System.Data;
using System.IO;
using System.Configuration;

namespace Iren.EpexDownloader
{
    class EpexDownloader
    {
        #region Variabili

        private WebClient _webClient = new WebClient();
        private HtmlAgilityPack.HtmlDocument _htmlDoc = new HtmlAgilityPack.HtmlDocument();
        private string _baseURL = "http://www.epexspot.com/en/market-data/dayaheadauction/auction-table/";
        private string _basePath = @"D:\Users\e-bergamin\Desktop";

        #endregion

        static void Main(string[] args)
        {
            EpexDownloader epexDwnloader = new EpexDownloader();

            DateTime dataInizio = new DateTime(2014, 8, 1);
            DateTime dataFine = DateTime.Now.AddDays(1);

            for (; dataInizio <= dataFine; dataInizio = dataInizio.AddDays(1))
            {
                Console.WriteLine("Data: " + dataInizio.ToString("dd/MM/yyyy"));
                epexDwnloader.Run(dataInizio);
            }
            Console.WriteLine("Done");
        }

        #region Costruttori

        public EpexDownloader()
        {
            _basePath = ConfigurationManager.AppSettings["basePath"] ?? _basePath;
            _baseURL = ConfigurationManager.AppSettings["baseURL"] ?? _baseURL;
        }

        #endregion

        #region Metodi

        public void Run(DateTime day)
        {
            bool is25hours = (day.Month == 10 && isLastSunday(day));
            bool is23hours = !is25hours && (day.Month == 3 && isLastSunday(day));

            string URL = _baseURL + day.ToString("yyyy-MM-dd") + "/FR";
            _htmlDoc.LoadHtml(_webClient.DownloadString(URL));

            //ottengo l'array delle date visualizzate
            HtmlNode dateRow = _htmlDoc.DocumentNode.SelectSingleNode("//div[@id='tab_fr']//table[@class='list hours responsive']//tr");
            List<DateTime> days = new List<DateTime>();
            foreach (HtmlNode col in dateRow.SelectNodes("th"))
            {
                DateTime d = new DateTime();
                if (DateTime.TryParseExact(col.InnerText + " " + day.Year, "ddd, MM/dd yyyy", new CultureInfo("en-US"), DateTimeStyles.None, out d))
                    days.Add(d);
            }

            Tuple<string, int>[] tabIDs = new Tuple<string, int>[] 
                { 
                    Tuple.Create("tab_fr", 987), 
                    Tuple.Create("tab_de", 924), 
                    Tuple.Create("tab_ch", 988)};

            foreach (Tuple<string, int> tabID in tabIDs)
            {
                HtmlNodeCollection tab = _htmlDoc.DocumentNode.SelectNodes("//div[@id='" + tabID.Item1 + "']//table[@class='list hours responsive']//tr[@class='no-border']");

                //la mia data ha 24 ore ma la tabella contiene anche la riga della 25-esima
                if (!is25hours && tab.Count() == 25)
                    tab.RemoveAt(3);

                DataTable dt = initTable();

                int i = 0;
                int index = days.IndexOf(day);
                foreach (HtmlNode row in tab)
                {
                    //seleziono il valore che mi interessa dalla tabella sapendo che index è 0-based e che le prime 2 colonne sono di intestazione
                    HtmlNode mgpVal = row.SelectSingleNode("td[" + (3 + index) + "]");
                    DataRow newRow = dt.NewRow();

                    newRow["Zona"] = tabID.Item2;
                    newRow["Data"] = day.ToString("yyyyMMdd") + (++i < 10 ? "0" : "") + i;
                    newRow["Mgp"] = 0;
                    decimal tmp;
                    if (Decimal.TryParse(mgpVal.InnerText.Replace('.', ','), out tmp))
                        newRow["MGP"] = tmp;

                    dt.Rows.Add(newRow);
                }

                //scrivo la tabella all'interno del caricatore
                string path = Path.Combine(_basePath, day.ToString("yyyyMMdd") + "_" + tabID.Item2 + ".xml");
                dt.WriteXml(path);
                dt.WriteXmlSchema(Path.Combine(_basePath, "schema.xml"));
            }
        }

        private DataTable initTable()
        {
            DataTable dt = new DataTable("Epex")
            {
                Columns =
                {
                    {"Zona", typeof(int)},
                    {"Data", typeof(string)},
                    {"Mgp", typeof(Decimal)}
                }
            };

            return dt;
        }

        private DateTime GetLastWeekdayOfMonth(DateTime date, DayOfWeek day)
        {
            DateTime lastDayOfMonth = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);
            int wantedDay = (int)day;
            int lastDay = (int)lastDayOfMonth.DayOfWeek;
            return lastDayOfMonth.AddDays(
                lastDay >= wantedDay ? wantedDay - lastDay : wantedDay - lastDay - 7);
        }

        private Boolean isLastSunday(DateTime date)
        {
            return date == GetLastWeekdayOfMonth(date, DayOfWeek.Sunday);
        }

        #endregion


    }
}
