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

        protected Dictionary<string, GotoObject> _definedGotos = new Dictionary<string, GotoObject>();

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
                where r["Sheet"].Equals(sheet)
                select r).ToDictionary(
                    r => r["Name"].ToString(), 
                    r => new GotoObject() 
                    { 
                        row = (int)r["Row"],
                        column = (int)r["Column"],
                        addressTo = r["AddressTo"].ToString()
                    }
                );
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
        public void AddName(int riga, params object[] parts)
        {
            _defNamesIndexByName.Add(GetName(parts), riga);
            _defNamesIndexByRow.Add(riga, GetName(parts));
        }
        public void AddGOTO(object siglaEntita, int row, int column, string addressTo = "")
        {
            GotoObject obj = new GotoObject() 
            {
                row = row,
                column = column,
                addressTo = addressTo
            };

            _definedGotos.Add(siglaEntita.ToString(), obj);
        }
        public void ChangeGOTOAddressTo(object siglaEntita, string addressTo)
        {
            _definedGotos[siglaEntita.ToString()].addressTo = addressTo;
        }

        public int GetColFromDate(string suffissoData, string suffissoOra = "H1")
        {
            string name = GetName(suffissoData, suffissoOra);
            return _defDatesIndexByName[name];
        }

        public int GetRowByName(params object[] parts)
        {
            return _defNamesIndexByName[GetName(parts)];
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
                        {"Sheet", typeof(string)},
                        {"Name", typeof(string)},
                        {"Row", typeof(int)},
                        {"Column", typeof(int)},
                        {"AddressTo", typeof(string)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["Sheet"], dt.Columns["Name"] };
            dt.TableName = name;
            return dt;
        }


        #endregion

        public void DumpToDataSet()
        {
            DataTable definedNames = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.NOMIDEFINITINEW];
            DataTable definedDates = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.DATEDEFINITE];
            DataTable definedGotos = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.GOTODEFINITI];

            ///////// nomi
            IEnumerable<DataRow> definedNamesRows =
                from r in definedNames.AsEnumerable()
                where r["Sheet"].Equals(_sheet)
                select r;

            foreach (var row in definedNamesRows)
            {
                if (_defNamesIndexByName.ContainsKey(row["Name"].ToString()))
                {
                    row["Row"] = _defNamesIndexByName[row["Name"].ToString()];
                    _defNamesIndexByName.Remove(row["Name"].ToString());
                }
            }

            if (_defNamesIndexByName.Count > 0) 
            {
                foreach(var ele in _defNamesIndexByName) 
                {
                    DataRow r = definedNames.NewRow();
                    r["Sheet"] = _sheet;
                    r["Name"] = ele.Key;
                    r["Row"] = ele.Value;
                    definedNames.Rows.Add(r);
                }
            }

            ///////// date
            IEnumerable<DataRow> definedDatesRows =
                from r in definedDates.AsEnumerable()
                where r["Sheet"].Equals(_sheet)
                select r;

            foreach (var row in definedDatesRows)
            {
                if (_defDatesIndexByName.ContainsKey(row["Name"].ToString()))
                {
                    row["Column"] = _defDatesIndexByName[row["Name"].ToString()];
                    _defDatesIndexByName.Remove(row["Name"].ToString());
                }
            }

            if (_defDatesIndexByName.Count > 0)
            {
                foreach (var ele in _defDatesIndexByName)
                {
                    DataRow r = definedDates.NewRow();
                    r["Sheet"] = _sheet;
                    r["Name"] = ele.Key;
                    r["Column"] = ele.Value;
                    definedDates.Rows.Add(r);
                }
            }

            ///////// goto
            IEnumerable<DataRow> definedGotosRows =
                from r in definedGotos.AsEnumerable()
                where r["Sheet"].Equals(_sheet)
                select r;

            foreach (var row in definedGotosRows)
            {
                if (_definedGotos.ContainsKey(row["Name"].ToString()))
                {
                    row["Row"] = _definedGotos[row["Name"].ToString()].row;
                    row["Column"] = _definedGotos[row["Name"].ToString()].column;
                    row["AddressTo"] = _definedGotos[row["Name"].ToString()].addressTo;
                    _definedGotos.Remove(row["Name"].ToString());
                }
            }

            if (_definedGotos.Count > 0)
            {
                foreach (var ele in _definedGotos)
                {
                    DataRow r = definedGotos.NewRow();
                    r["Sheet"] = _sheet;
                    r["Name"] = ele.Key;
                    r["Row"] = ele.Value.row;
                    r["Column"] = ele.Value.column;
                    r["AddressTo"] = ele.Value.addressTo;
                    definedGotos.Rows.Add(r);
                }
            }
        }
    }

    #region Classi supporto

    public class GotoObject
    {
        public int row, column;
        public string addressTo;
    }

    #endregion
}
