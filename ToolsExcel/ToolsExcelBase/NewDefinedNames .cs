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
        List<string> _days;

        protected Dictionary<string, int> _defDatesIndexByName = new Dictionary<string,int>();
        protected Dictionary<int, string> _defDatesIndexByCol = new Dictionary<int, string>();

        protected Dictionary<string, int> _defNamesIndexByName = new Dictionary<string, int>();
        protected ILookup<int, string> _defNamesIndexByRow;

        protected Dictionary<string, GotoObject> _definedGotos = new Dictionary<string, GotoObject>();

        #endregion

        #region Proprietà

        public string[] DaySuffx
        {
            get
            {
                return _days.ToArray();
            }
        }

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

            _defNamesIndexByName = names.ToDictionary(r => r["Name"].ToString(), r => (int)r["Row"]);
            _defNamesIndexByRow = names.ToLookup(r => (int)r["Row"], r => r["Name"].ToString());

            IEnumerable<DataRow> dates =
                from DataRow r in definedDates.AsEnumerable()
                where r["Sheet"].Equals(sheet)
                select r;

            DataView distinctDays = new DataView(definedDates);
            distinctDays.RowFilter = "Sheet = '" + _sheet + "'";
            _days = 
                (from r in distinctDays.ToTable(true, "Date").AsEnumerable()
                 select r["Date"].ToString()).ToList();

            _defDatesIndexByName = dates.ToDictionary(r => GetName(r["Date"].ToString(), r["Hour"].ToString()), r => (int)r["Column"]);
            _defDatesIndexByCol = dates.ToDictionary(r => (int)r["Column"], r => GetName(r["Date"].ToString(), r["Hour"].ToString()));

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

                string suffissoData = Date.GetSuffissoData(giorno);
                for (int ora = 0; ora < oreGiorno; ora++)
                {
                    string data = GetName(suffissoData, Date.GetSuffissoOra(ora + 1));
                    _defDatesIndexByName.Add(data, colStart);
                    _defDatesIndexByCol.Add(colStart, data);
                    colStart++;
                }
                _days.Add(suffissoData);
            }
        }
        public void AddName(int riga, params object[] parts)
        {
            _defNamesIndexByName.Add(GetName(parts), riga);
            //_defNamesIndexByRow(riga, GetName(parts));
        }
        public void AddDate(int col, params object[] parts)
        {
            _defDatesIndexByName.Add(GetName(parts), col);
            _defDatesIndexByCol.Add(col, GetName(parts));
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

        public int GetFirstCol()
        {
            return _defDatesIndexByCol.ElementAt(0).Key;
        }
        public int GetFirstRow()
        {
            return _defNamesIndexByName.ElementAt(0).Value;
        }
        public int GetColFromDate()
        {
            return GetColFromDate(Date.GetSuffissoData(DataBase.DataAttiva));
        }
        public int GetColFromDate(DateTime giorno)
        {
            return GetColFromDate(Date.GetSuffissoData(giorno));
        }
        public int GetColFromDate(string suffissoData, string suffissoOra = "H1")
        {
            if (Struct.tipoVisualizzazione == "V")
                suffissoData = Date.GetSuffissoData(DataBase.DataAttiva);

            string name = GetName(suffissoData, suffissoOra);
            return _defDatesIndexByName[name];
        }
        public int GetColFromName(params object[] parts)
        {
            return _defDatesIndexByName[GetName(parts)];
        }
        public int GetColOffset()
        {
            if (Struct.tipoVisualizzazione == "O")
                return _defDatesIndexByName.Count;

            return 25;
        }
        public int GetColOffsetRiepilogo()
        {
            return _defDatesIndexByName.Count;
        }
        public int GetRowOffset()
        {
            return _defNamesIndexByName.Count;
        }
        public int GetColOffset(DateTime data)
        {
            return GetColOffset(Date.GetSuffissoData(data));
        }
        public int GetColOffset(string suffissoData)
        {
            var date =
                from kv in _defDatesIndexByName
                where kv.Key.Substring(0, suffissoData.Length).CompareTo(suffissoData) <= 0
                select kv;

            return date.Count();
        }
        public int GetDayOffset(string suffissoData)
        {
            if (Struct.tipoVisualizzazione == "V")
                return 25;

            var date =
                from kv in _defDatesIndexByName
                where kv.Key.StartsWith(suffissoData)
                select kv;

            return date.Count();
        }
        public int GetDayOffset(DateTime giorno)
        {
            return GetDayOffset(Date.GetSuffissoData(giorno));
        }

        public int GetRowByName(params object[] parts)
        {
            return _defNamesIndexByName[GetName(parts)];
        }
        public int GetRowByName(string name)
        {
            return _defNamesIndexByName[name];
        }
        public int GetRowByName(object siglaEntita, object siglaInformazione, string suffissoData)
        {
            string name = GetName(siglaEntita, siglaInformazione, Struct.tipoVisualizzazione == "O" ? "" : suffissoData);
            return GetRowByName(name);
        }

        public Range Get(object siglaEntita, object siglaInformazione)
        {
            int row = GetRowByName(siglaEntita, siglaInformazione);
            int col = GetFirstCol();

            return new Range(row, col);
        }
        public Range Get(object siglaEntita, object siglaInformazione, string suffissoData)
        {
            return Get(siglaEntita, siglaInformazione, suffissoData, "H1");
        }
        public Range Get(object siglaEntita, object siglaInformazione, string suffissoData, string suffissoOra)
        {
            string name = GetName(siglaEntita, siglaInformazione, Struct.tipoVisualizzazione == "O" ? "" : suffissoData);
            //string suffData = Struct.tipoVisualizzazione == "O" ? suffissoData : Date.GetSuffissoData(DataBase.DataAttiva);

            int row = GetRowByName(name);
            int col = GetColFromDate(suffissoData, suffissoOra);

            return new Range(row, col);
        }


        #endregion

        #region Metodi Statici

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
        public static DataTable GetDefaultDateTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(String)},
                        {"Date", typeof(String)},
                        {"Hour", typeof(String)},
                        {"Column", typeof(int)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["Sheet"], dt.Columns["Date"], dt.Columns["Hour"] };
            dt.TableName = name;
            return dt;
        }
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
                    r["Date"] = ele.Key.Split(Simboli.UNION[0])[0];
                    r["Hour"] = ele.Key.Split(Simboli.UNION[0])[1];
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
