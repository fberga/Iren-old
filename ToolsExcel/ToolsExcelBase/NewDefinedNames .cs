using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using Iren.ToolsExcel.Utility;
using System.Collections;

namespace Iren.ToolsExcel.Base
{
    public class NewDefinedNames
    {
        #region Variabili
        
        string _sheet;

        protected Dictionary<string, int> _defDatesIndexByName = new Dictionary<string,int>();
        protected Dictionary<int, string> _defDatesIndexByCol = new Dictionary<int, string>();

        protected Dictionary<string, int> _defNamesIndexByName = new Dictionary<string, int>();
        protected Dictionary<int, string> _defNamesIndexByRow = new Dictionary<int, string>();

        protected List<string> _definedGotos = new List<string>();

        #endregion

        #region Costruttori

        public NewDefinedNames(string sheet)
        {
            _sheet = sheet;

            DataTable definedNames = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.NOMIDEFINITINEW];
            DataTable definedDates = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.DATEDEFINITE];
            DataTable definedGotos = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.GOTODEFINITI];

            IEnumerable<DataRow> names =
                from DataRow r in definedNames.AsEnumerable()
                where r["Sheet"].Equals(sheet)
                select r;

            _defDatesIndexByName = names.ToDictionary(r => r["Name"].ToString(), r => (int)r["Row"]);
            _defNamesIndexByRow = names.ToDictionary(r => (int)r["Row"], r => r["Name"].ToString());

            IEnumerable<DataRow> dates =
                from DataRow r in definedDates.AsEnumerable()
                where r["Sheet"].Equals(sheet)
                select r;

            _defDatesIndexByName = names.ToDictionary(r => r["Name"].ToString(), r => (int)r["Column"]);
            _defDatesIndexByCol = names.ToDictionary(r => (int)r["Column"], r => r["Name"].ToString());

            _definedGotos =
                (from DataRow r in definedGotos.AsEnumerable()
                select r["Name"].ToString()).ToList();
        }

        #endregion

        #region Metodi

        public void DefineDates(DateTime dataInizio, DateTime dataFine, int colStart, bool data0H24)
        {
            if (data0H24)
            {
                string data = GetName(Date.GetSuffissoData(dataInizio.AddDays(-1)), Date.GetSuffissoOra(24));
                _defDatesIndexByName.Add(data, colStart);
                _defDatesIndexByCol.Add(colStart, data);
                colStart++;
            }

            for (DateTime giorno = dataInizio; giorno <= dataFine; giorno = giorno.AddDays(1))
            {
                int oreGiorno = Struct.tipoVisualizzazione == "O" ? Date.GetOreGiorno(giorno) : 25;

                for (int ora = 0; ora < oreGiorno; ora++)
                {
                    string data = GetName(Date.GetSuffissoData(giorno), Date.GetSuffissoOra(ora + 1));
                    _defDatesIndexByName.Add(data, colStart);
                    _defDatesIndexByCol.Add(colStart, data);
                    colStart++;
                }
            }
        }
        public void Add(object siglaEntita, object siglaInfo, int riga)
        {
            _defNamesIndexByName.Add(GetName(siglaEntita, siglaInfo), riga);
            _defNamesIndexByRow.Add(riga, GetName(siglaEntita, siglaInfo));
        }
        
        public void AddGOTO(object siglaEntita)
        {
            _definedGotos.Add(GetName(siglaEntita, siglaInfo), riga);
            _defNamesIndexByRow.Add(riga, GetName(siglaEntita, siglaInfo));
        }

        #endregion

        #region Metodi Statici

        /// <summary>
        /// Da una lista di oggetti in input, compone il nome con il simbolo di unione.
        /// </summary>
        /// <param name="parts">Lista di stringhe che andranno a comporre il nome in output</param>
        /// <returns>Restituisce la stringa che rappresenta il nome</returns>
        public static string GetName(params object[] parts)
        {
            string o = "";
            bool first = true;
            foreach (object part in parts)
            {
                if (part != null && part.ToString() != "")
                {
                    o += (!first ? Simboli.UNION : "") + part;
                    first = false;
                }
            }
            return o;
        }
        /// <summary>
        /// Inizializza la tabella dei nomi assegnandole un nome e la restituisce.
        /// </summary>
        /// <param name="name">Il nome da assegnare alla tabella per la serializzazione.</param>
        /// <returns>Ritorna una nuova istanza della tabella dei nomi.</returns>
        public static DataTable GetDefaultNameTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(String)},
                        {"Name", typeof(String)},
                        {"Row", typeof(int)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["Sheet"], dt.Columns["Name"] };
            dt.TableName = name;
            return dt;
        }
        /// <summary>
        /// Inizializza la tabella delle date definite assegnandole un nome e la restituisce.
        /// </summary>
        /// <param name="name">Il nome da assegnare alla tabella per la serializzazione.</param>
        /// <returns>Ritorna una nuova istanza della tabella dei nomi.</returns>
        public static DataTable GetDefaultDateTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(String)},
                        {"Name", typeof(String)},
                        {"Column", typeof(int)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["Sheet"], dt.Columns["Column"] };
            dt.TableName = name;
            return dt;
        }
        /// <summary>
        /// Inizializza la tabella delle date definite assegnandole un nome e la restituisce.
        /// </summary>
        /// <param name="name">Il nome da assegnare alla tabella per la serializzazione.</param>
        /// <returns>Ritorna una nuova istanza della tabella dei nomi.</returns>
        public static DataTable GetDefaultGOTOTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(String)},
                        {"Name", typeof(String)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["Sheet"], dt.Columns["Column"] };
            dt.TableName = name;
            return dt;
        }


        #endregion
    }
}
