using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Iren.FrontOffice.Base
{
    public class DefinedNames
    {
        #region Variabili

        DataTable _definedNames;
        DataView _definedNamesView;

        #endregion

        #region Costruttori

        public DefinedNames()
        {
            _definedNames = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.NOMIDEFINITI];
            _definedNamesView = new DataView(_definedNames);
        }

        ~DefinedNames()
        {
            _definedNames.Dispose();
            _definedNamesView.Dispose();
        }

        #endregion

        #region Overload Operatori

        public Tuple<int, int> this[string key]
        {
            get
            {
                return GetFirstCell(key);
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

        public bool IsRange(string name)
        {
            _definedNamesView.RowFilter = "Nome='" + name + "'";
            if (_definedNamesView.Count == 0)
                return false;
            return _definedNamesView[0]["R1"] != _definedNamesView[0]["R2"] || _definedNamesView[0]["C1"] != _definedNamesView[0]["C2"];
        }        

        public string[] Get(int row, int column)
        {
            _definedNamesView.RowFilter = "R1=" + row + " AND C1=" + column;
            
            if (_definedNamesView.Count == 0)
                return null;

            List<string> o = new List<string>();

            foreach (DataRowView name in _definedNamesView)
                o.Add(name["Nome"].ToString());

            return o.ToArray();
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
    }
}
