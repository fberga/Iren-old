using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;

namespace Iren.ToolsExcel.Base
{
    public class DefinedNames
    {
        #region Variabili

        protected DataTable _definedNames;
        protected DataView _definedNamesView;
        protected string _foglio;
        //protected int _row = -1;
        //protected int _column = -1;
        //protected string _name = null;
        //protected string _exclude = "";
        //protected bool _excludeDATA0H24 = false;

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

        public void Add(string nome, Tuple<int, int> cella1, Tuple<int, int> cella2 = null, bool editabile = false, bool salvaDB = false, bool annotaModifica = false)
        {
            DataRow r = _definedNames.NewRow();
            cella2 = cella2 ?? cella1;
            r["Foglio"] = _foglio;
            r["Nome"] = nome;
            r["R1"] = cella1.Item1;
            r["C1"] = cella1.Item2;
            r["R2"] = cella2.Item1;
            r["C2"] = cella2.Item2;
            r["Editabile"] = editabile;
            r["SalvaDB"] = salvaDB;
            r["AnnotaModifica"] = annotaModifica;

            _definedNames.Rows.Add(r);
        }
        public void Add(string name, int row, int column, bool editabile = false, bool salvaDB = false, bool annotaModifica = false)
        {
            Add(name, Tuple.Create(row, column), editabile: editabile, salvaDB: salvaDB, annotaModifica: annotaModifica);
        }
        public void Add(string name, int row1, int column1, int row2, int column2, bool editabile = false, bool salvaDB = false, bool annotaModifica = false)
        {
            Add(name, Tuple.Create(row1, column1), Tuple.Create(row2, column2), editabile, salvaDB, annotaModifica);
        }


        public string[] Get(int row, int column)
        {
            string filter = "Foglio = '" + _foglio + "' AND R1 = " + row + " AND C1 = " + column;
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;
            
            if (_definedNamesView.Count == 0)
                return null;

            List<string> o = new List<string>();

            foreach (DataRowView name in _definedNamesView)
                o.Add(name["Nome"].ToString());

            return o.ToArray();
        }
        public Tuple<int, int>[] Get(string name, bool excludeDATA0H24 = false)
        {
            return Get(name, "", excludeDATA0H24);
        }
        public Tuple<int, int>[] Get(string name, string exclude, bool excludeDATA0H24 = false)
        {
            string filter = "";
            name = PrepareName(name);

            if (!excludeDATA0H24)
                filter = "Nome LIKE '" + name + "%'";
            else
                filter = "Nome LIKE '" + name + "%' AND Nome NOT LIKE '%DATA0.H24%'";

            if (exclude != "")
                filter += " AND Nome NOT LIKE '%" + exclude + "%'";

            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

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

        public bool IsDefined(string name)
        {
            name = PrepareName(name);
            string filter = "Foglio = '" + _foglio + "' AND Nome LIKE '" + name + "%'";

            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            return _definedNamesView.Count > 0;
        }
        public bool IsDefined(int row, int column)
        {
            string filter = "Foglio = '" + _foglio + "' AND R1 = " + row + " AND C1 = " + column;
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            return _definedNamesView.Count > 0;
        }

        public bool IsRange(string name)
        {
            string filter = "Nome='" + name + "'";
            
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            if (_definedNamesView.Count == 0)
                return false;
            return _definedNamesView[0]["R1"] != _definedNamesView[0]["R2"] || _definedNamesView[0]["C1"] != _definedNamesView[0]["C2"];
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

        public bool Editabile(string name)
        {
            name = PrepareName(name);
            string filter = "Foglio = '" + _foglio + "' AND Nome LIKE '" + name + "%'";
            if (_definedNamesView.RowFilter != filter)            
                _definedNamesView.RowFilter = filter;
            
            if (_definedNamesView.Count == 0)
                return false;

            return (bool)_definedNamesView[0]["Editabile"];
        }
        public bool Editabile(int row, int column)
        {
            string filter = "Foglio = '" + _foglio + "' AND R1 = " + row + " AND C1 = " + column;
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;
            
            if (_definedNamesView.Count == 0)
                return false;

            return (bool)_definedNamesView[0]["Editabile"];
        }
        
        public bool SalvaDB(string name)
        {
            name = PrepareName(name);
            string filter = "Foglio = '" + _foglio + "' AND Nome LIKE '" + name + "%'";
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;
            
            if (_definedNamesView.Count == 0)
                return false;

            return (bool)_definedNamesView[0]["SalvaDB"];
        }
        public bool SalvaDB(int row, int column)
        {
            string filter = "Foglio = '" + _foglio + "' AND R1 = " + row + " AND C1 = " + column;
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            if (_definedNamesView.Count == 0)
                return false;

            return (bool)_definedNamesView[0]["SalvaDB"];
        }
        
        public bool AnnotaModifica(string name)
        {
            name = PrepareName(name);
            string filter = "Foglio = '" + _foglio + "' AND Nome LIKE '" + name + "%'";
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            if (_definedNamesView.Count == 0)
                return false;

            return (bool)_definedNamesView[0]["AnnotaModifica"];
        }
        public bool AnnotaModifica(int row, int column)
        {
            string filter = "Foglio = '" + _foglio + "' AND R1 = " + row + " AND C1 = " + column;
            if (_definedNamesView.RowFilter != filter)
                _definedNamesView.RowFilter = filter;

            if (_definedNamesView.Count == 0)
                return false;

            return (bool)_definedNamesView[0]["AnnotaModifica"];
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
                        {"C2", typeof(int)},
                        {"Editabile", typeof(bool)},
                        {"SalvaDB", typeof(bool)},
                        {"AnnotaModifica", typeof(bool)}
                    }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["Foglio"], dt.Columns["Nome"] };
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
        public static bool IsDefined(string sheetName, int row, int column)
        {
            DataView definedNamesView = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.NOMIDEFINITI].DefaultView;
            definedNamesView.RowFilter = "Foglio = '" + sheetName + "' AND R1 = " + row + " AND C1 = " + column;

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
