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

        protected Dictionary<string, object> _addressFrom = new Dictionary<string, object>();
        protected Dictionary<object, string> _addressTo = new Dictionary<object, string>();

        public enum InitType
        {
            All, AllThisSheet, OnlyNaming, OnlyGOTOs, OnlyGOTOsThisSheet
        }

        #endregion

        #region Proprietà

        public string[] DaySuffx
        {
            get
            {
                return _days.ToArray();
            }
        }
        public string Sheet
        {
            get { return _sheet; }
        }

        #endregion

        #region Costruttori

        private void InitNaming()
        {
            DataTable definedNames = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.NOMIDEFINITINEW];
            DataTable definedDates = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.DATEDEFINITE];


            IEnumerable<DataRow> names =
                from DataRow r in definedNames.AsEnumerable()
                where r["Sheet"].Equals(_sheet)
                select r;

            _defNamesIndexByName = names.ToDictionary(r => r["Name"].ToString(), r => (int)r["Row"]);
            _defNamesIndexByRow = names.ToLookup(r => (int)r["Row"], r => r["Name"].ToString());

            IEnumerable<DataRow> dates =
                from DataRow r in definedDates.AsEnumerable()
                where r["Sheet"].Equals(_sheet)
                select r;

            DataView distinctDays = new DataView(definedDates);
            distinctDays.RowFilter = "Sheet = '" + _sheet + "'";
            _days =
                (from r in distinctDays.ToTable(true, "Date").AsEnumerable()
                 select r["Date"].ToString()).ToList();

            _defDatesIndexByName = dates.ToDictionary(r => GetName(r["Date"].ToString(), r["Hour"].ToString()), r => (int)r["Column"]);
            _defDatesIndexByCol = dates.ToDictionary(r => (int)r["Column"], r => GetName(r["Date"].ToString(), r["Hour"].ToString()));
        }
        private void InitGOTOs(bool thisSheet = false)
        {
            DataTable addressFromTable = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.ADDRESSFROM];
            DataTable addressToTable = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.ADDRESSTO];

            _addressFrom =
               (from DataRow r in addressFromTable.AsEnumerable()
                where !thisSheet || r["Sheet"].Equals(_sheet)
                select r).ToDictionary(
                    r => r["AddressFrom"].ToString(),
                    r => r["SiglaEntita"]
                );

            _addressTo =
               (from DataRow r in addressToTable.AsEnumerable()
                where !thisSheet || r["Sheet"].Equals(_sheet)
                select r).ToDictionary(
                    r => r["SiglaEntita"],
                    r => r["AddressTo"].ToString()
                );
        }

        public NewDefinedNames(string sheet, InitType type = InitType.OnlyNaming)
        {
            _sheet = sheet;

            switch (type)
            {
                case InitType.All:
                    InitNaming();
                    InitGOTOs();
                    break;
                case InitType.OnlyNaming:
                    InitNaming();
                    break;
                case InitType.OnlyGOTOs:
                    InitGOTOs();
                    break;
                case InitType.OnlyGOTOsThisSheet:
                    InitGOTOs(true);
                    break;
            }
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
        public void AddGOTO(object siglaEntita, string addressFrom)
        {
            _addressFrom.Add("'" + _sheet + "'!" + addressFrom, siglaEntita);
        }
        public void AddGOTO(object siglaEntita, string addressFrom, string addressTo)
        {
            AddGOTO(siglaEntita, addressFrom);
            _addressTo.Add(siglaEntita, "'" + _sheet + "'!" + addressTo);
        }
        public void ChangeGOTOAddressTo(object siglaEntita, string addressTo)
        {
            _addressTo[siglaEntita] = "'" + _sheet + "'!" + addressTo;
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

        public Range Get(params object[] parts)
        {
            if (parts.Length == 2)
                return new Range(GetRowByName(GetName(parts)), GetFirstCol());

            List<string> nameParts = new List<string>();
            List<string> dateParts = new List<string>();
            bool date = false;
            foreach (var part in parts)
            {
                date = date || Regex.IsMatch(part.ToString(), @"DATA\d+");
                if (!date)
                    nameParts.Add(part.ToString());
                else
                    dateParts.Add(part.ToString());
            }

            if (!date)
                return new Range(GetRowByName(GetName(nameParts)), GetFirstCol());

            if(dateParts[0].Contains(Simboli.UNION)) 
            {
                string[] suffissoDataOra = dateParts[0].Split(Simboli.UNION[0]);
                dateParts = new List<string>() { suffissoDataOra[0], suffissoDataOra[1] };
            }
            else if (dateParts.Count == 1)
                dateParts.Add(Date.GetSuffissoOra(1));
            
            int row = GetRowByName(GetName(nameParts, Struct.tipoVisualizzazione == "O" ? "" : dateParts[0]));
            int col = GetColFromDate(dateParts[0], dateParts[1]);

            return new Range(row, col);

        }
        //public Range Get(object siglaEntita, object siglaInformazione)
        //{
        //    int row = GetRowByName(siglaEntita, siglaInformazione);
        //    int col = GetFirstCol();

        //    return new Range(row, col);
        //}
        //public Range Get(object siglaEntita, object siglaInformazione, string suffissoData)
        //{
        //    return Get(siglaEntita, siglaInformazione, suffissoData, "H1");
        //}
        //public Range Get(object siglaEntita, object siglaInformazione, string suffissoData, string suffissoOra)
        //{
        //    string name = GetName(siglaEntita, siglaInformazione, Struct.tipoVisualizzazione == "O" ? "" : suffissoData);
        //    int row = GetRowByName(name);
        //    int col = GetColFromDate(suffissoData, suffissoOra);

        //    return new Range(row, col);
        //}

        public string GetGOTO(string addressFrom)
        {
            if (_addressFrom.ContainsKey("'" + _sheet + "'!" + addressFrom))
                return GetGOTO(_addressFrom["'" + _sheet + "'!" + addressFrom]);

            return "";
        }
        public string GetGOTO(object siglaEntita)
        {
            if (_addressTo.ContainsKey(siglaEntita))
                return _addressTo[siglaEntita];

            return "";
        }
        public string GetAddressFromGOTO(int i)
        {
            return _addressFrom.ElementAt(i).Key;
        }

        #endregion

        #region Metodi Statici

        public static string GetName(List<string> parts, string name)
        {            
            parts.Add(name);
            return GetName(parts);
        }
        public static string GetName(string name, List<string> parts)
        {
            List<string> list = new List<string>();
            list.Add(name);
            list.AddRange(parts);

            return GetName(list);
        }
        public static string GetName(params List<string>[] parts)
        {
            string o = "";
            bool first = true;
            foreach (List<string> part in parts)
            {
                foreach (string part1 in part)
                {
                    if(part1 != null && part1 != "")
                    {
                        o += (!first ? Simboli.UNION : "") + part1;
                        first = false;
                    }
                }
            }
            return o;
        }
        public static string GetName(params object[] parts)
        {
            List<string> list = new List<string>();
            foreach (object part in parts)
                if(part.GetType() == typeof(string))
                    list.Add(part.ToString());

            return GetName(list);

            //string o = "";
            //bool first = true;
            //foreach (object part in parts)
            //{
            //    if (part != null && part.ToString() != "")
            //    {
            //        o += (!first ? Simboli.UNION : "") + part;
            //        first = false;
            //    }
            //}
            //return o;
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
        public static DataTable GetDefaultAddressFromTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(string)},
                        {"AddressFrom", typeof(string)},
                        {"SiglaEntita", typeof(string)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["AddressFrom"] };
            dt.TableName = name;
            return dt;
        }
        public static DataTable GetDefaultAddressToTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Sheet", typeof(string)},
                        {"SiglaEntita", typeof(string)},
                        {"AddressTo", typeof(string)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["SiglaEntita"] };
            dt.TableName = name;
            return dt;
        }

        public static string GetSheetName(object siglaEntita)
        {
            DataTable dt = DataBase.LocalDB.Tables[DataBase.Tab.NOMIDEFINITINEW];

            string s =
                (from r in dt.AsEnumerable()
                 where r["Name"].ToString().Contains(siglaEntita.ToString()) && !r["Sheet"].Equals("Main")
                 select r["Sheet"].ToString()).First();

            return s ?? "";
        }

        #endregion

        public void DumpToDataSet()
        {
            DataTable definedNames = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.NOMIDEFINITINEW];
            DataTable definedDates = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.DATEDEFINITE];
            DataTable addressFromTable = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.ADDRESSFROM];
            DataTable addressToTable = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.ADDRESSTO];

            ///////// nomi
            //IEnumerable<DataRow> definedNamesRows =
            //    from r in definedNames.AsEnumerable()
            //    where r["Sheet"].Equals(_sheet)
            //    select r;

            //foreach (var row in definedNamesRows)
            //{
            //    if (_defNamesIndexByName.ContainsKey(row["Name"].ToString()))
            //    {
            //        row["Row"] = _defNamesIndexByName[row["Name"].ToString()];
            //        _defNamesIndexByName.Remove(row["Name"].ToString());
            //    }
            //}

            //if (_defNamesIndexByName.Count > 0) 
            //{
                foreach(var ele in _defNamesIndexByName) 
                {
                    DataRow r = definedNames.NewRow();
                    r["Sheet"] = _sheet;
                    r["Name"] = ele.Key;
                    r["Row"] = ele.Value;
                    definedNames.Rows.Add(r);
                }
            //}

            ///////// date
            //IEnumerable<DataRow> definedDatesRows =
            //    from r in definedDates.AsEnumerable()
            //    where r["Sheet"].Equals(_sheet)
            //    select r;

            //foreach (var row in definedDatesRows)
            //{
            //    if (_defDatesIndexByName.ContainsKey(row["Name"].ToString()))
            //    {
            //        row["Column"] = _defDatesIndexByName[row["Name"].ToString()];
            //        _defDatesIndexByName.Remove(row["Name"].ToString());
            //    }
            //}

            //if (_defDatesIndexByName.Count > 0)
            //{
                foreach (var ele in _defDatesIndexByName)
                {
                    string[] dateTime = ele.Key.Split(Simboli.UNION[0]);

                    DataRow r = definedDates.NewRow();
                    r["Sheet"] = _sheet;
                    r["Date"] = dateTime[0];
                    r["Hour"] = dateTime.Length == 2 ? ele.Key.Split(Simboli.UNION[0])[1] : "";
                    r["Column"] = ele.Value;
                    definedDates.Rows.Add(r);
                }
            //}

            ///////// goto
            //IEnumerable<KeyValuePair<string, object>> defAddressFrom =
            //    from kv in _addressFrom
            //    where kv.Key.StartsWith("'" + _sheet + "'!")
            //    select kv;
            foreach (var ele in _addressFrom)
            {
                DataRow r = addressFromTable.NewRow();
                r["Sheet"] = _sheet;
                r["AddressFrom"] = ele.Key;
                r["SiglaEntita"] = ele.Value;
                addressFromTable.Rows.Add(r);
            }
            //IEnumerable<KeyValuePair<object, string>> defAddressTo =
            //    from kv in _addressTo
            //    where kv.Value.StartsWith("'" + _sheet + "'!")
            //    select kv;
            foreach (var ele in _addressTo)
            {
                DataRow r = addressToTable.NewRow();
                r["Sheet"] = _sheet;
                r["SiglaEntita"] = ele.Key;
                r["AddressTo"] = ele.Value;
                addressToTable.Rows.Add(r);
            }
        }
    }

    #region Classi supporto

    //public class GotoObject : IEqualityComparer
    //{
    //    private string _sheet;
    //    private int _row, _column;
    //    private string _address;
    //    private string _addressTo;

    //    public string Sheet
    //    {
    //        get { return _sheet; }
    //    }

    //    public int Row
    //    {
    //        get { return _row; }
    //    }

    //    public int Column
    //    {
    //        get { return _column; }
    //    }

    //    public string Address
    //    {
    //        get { return _address; }
    //    }

    //    public string AddressTo
    //    {
    //        get { return _addressTo; }
    //    }

    //    public GotoObject(string sheet, int row, int column)
    //    {
    //        _sheet = sheet;
    //        _row = row;
    //        _column = column;
    //        _address = "'" + _sheet + "'!" + Range.GetRange(_row, _column);
    //    }
    //    public GotoObject(string sheet, int row, int column, string addressTo) 
    //        : this(sheet, row, column)
    //    {
    //        _addressTo = addressTo;
    //    }

    //    public bool Equals(object x, object y)
    //    {
    //        GotoObject obj1 = (GotoObject)x;
    //        GotoObject obj2 = (GotoObject)y;

    //        return obj1.Address == obj2.Address;
    //    }

    //    public int GetHashCode(object obj)
    //    {
    //        GotoObject obj1 = (GotoObject)obj;
    //        return obj1.Address.GetHashCode();
    //    }
    //}

    #endregion
}
