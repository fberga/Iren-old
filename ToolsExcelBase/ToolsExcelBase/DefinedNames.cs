using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Iren.FrontOffice.Base
{
    public class DefinedNames
    {
        #region Variabili

        DataTable _definedNames;
        DataView _definedNamesView;
        string _foglio;

        #endregion

        #region Costruttori

        public DefinedNames(string foglio)
        {
            _foglio = foglio;
            _definedNames = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.NOMIDEFINITI];
            _definedNamesView = new DataView(_definedNames);
        }

        #endregion

        #region Overload Operatori

        public Tuple<int, int>[] this[string key]
        {
            get
            {
                return Get(key);
            }
        }

        public string[] this[int r1, int c1]
        {
            get
            {
                return Get(r1, c1);
            }
        }

        #endregion

        #region Metodi

        public void Add(string nome, Tuple<int, int> cella1, Tuple<int, int> cella2 = null)
        {
            DataRow r = _definedNames.NewRow();
            cella2 = cella2 ?? cella1;
            r["Foglio"] = _foglio;
            r["Nome"] = nome;
            r["R1"] = cella1.Item1;
            r["C1"] = cella1.Item2;
            r["R2"] = cella2.Item1;
            r["C2"] = cella2.Item2;

            //TODO controllare se nome esiste già

            _definedNames.Rows.Add(r);
        }
        public void Add(string name, int row, int column)
        {
            Add(name, Tuple.Create(row, column));
        }
        public void Add(string name, int row1, int column1, int row2, int column2)
        {
            Add(name, Tuple.Create(row1, column1), Tuple.Create(row2, column2));
        }

        public bool IsRange(string name)
        {
            _definedNamesView.RowFilter = "Nome='" + name + "'";
            if (_definedNamesView.Count == 0)
                return false;
            return _definedNamesView[0]["R1"] != _definedNamesView[0]["R2"] || _definedNamesView[0]["C1"] != _definedNamesView[0]["C2"];
        }        

        public string[] Get(int row, int column)
        {
            _definedNamesView.RowFilter = "Foglio='" + _foglio + "' AND R1=" + row + " AND C1=" + column;
            
            if (_definedNamesView.Count == 0)
                return null;

            List<string> o = new List<string>();

            foreach (DataRowView name in _definedNamesView)
                o.Add(name["Nome"].ToString());

            return o.ToArray();
        }
        public Tuple<int, int>[] Get(string name)
        {
            name = PrepareName(name);

            _definedNamesView.RowFilter = "Nome LIKE '" + name + "%'";

            if (_definedNamesView.Count == 0)
                return null;

            Tuple<int, int>[] o = new Tuple<int, int>[_definedNamesView.Count];
            int i = 0;
            foreach (DataRowView defName in _definedNamesView)
            {
                o[i++] = Tuple.Create(int.Parse(defName["R1"].ToString()), int.Parse(defName["C1"].ToString()));
            }

            return o;
        }
        public Tuple<int,int> GetFirstCell(string name)
        {
            _definedNamesView.RowFilter = "Nome='" + name + "'";

            if (_definedNamesView.Count == 0)
                return null;

            return Tuple.Create(int.Parse(_definedNamesView[0]["R1"].ToString()), int.Parse(_definedNamesView[0]["C1"].ToString()));
        }

        public Tuple<int, int>[] GetRange(string name)
        {
            if (!IsRange(name))
                return null;

            if (_definedNamesView.Count == 0)
                return null;

            return new Tuple<int, int>[2] 
                { 
                    Tuple.Create(int.Parse(_definedNamesView[0]["R1"].ToString()), int.Parse(_definedNamesView[0]["C1"].ToString())),
                    Tuple.Create(int.Parse(_definedNamesView[0]["R2"].ToString()), int.Parse(_definedNamesView[0]["C2"].ToString()))
                };
        }

        #endregion

        #region Metodi Statici

        private static string PrepareName(string name)
        {
            //se il nome non fa parte del riepilogo e non finisce con il suffisso data ora, aggiungo un punto
            if (!Regex.IsMatch(name, @"GRAFICO\d+|RIEPILOGO|DATA\d+\.H\d+|\.T\."))
                name += Simboli.UNION;
            return name;
        }

        public static DataTable GetDefaultTable(string name)
        {
            DataTable dt = new DataTable()
            {
                Columns =
                    {
                        {"Foglio", typeof(String)},
                        {"Nome", typeof(String)},
                        {"R1", typeof(int)},
                        {"C1", typeof(int)},
                        {"R2", typeof(int)},
                        {"C2", typeof(int)}
                    }
            };
            dt.TableName = name;
            return dt;
        }

        public static string GetSheetName(object name)
        {
            DataView definedNamesView = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.NOMIDEFINITI].DefaultView;
            definedNamesView.RowFilter = "Nome LIKE'" + name + "%'";
            if (definedNamesView.Count == 0)
                return null;

            return definedNamesView[0]["Foglio"].ToString();
        }

        public static bool IsDefined(string sheetName, string cellName)
        {
            cellName = PrepareName(cellName);
            DataView definedNamesView = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.NOMIDEFINITI].DefaultView;
            definedNamesView.RowFilter = "Foglio = '" + sheetName + "' AND Nome LIKE '" + cellName + "%'";

            return definedNamesView.Count > 0;
        }

        public static string GetName(params object[] parts)
        {
            string o = "";
            bool first = true;
            foreach (object part in parts)
            {
                if (part != null)
                {
                    o += (!first && part != "" ? Simboli.UNION : "") + part;
                    first = false;
                }
            }
            return o;
        }

        #endregion
    }
}
