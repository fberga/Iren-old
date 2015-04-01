using System;
using System.Text.RegularExpressions;

namespace Iren.ToolsExcel.Base
{
    public class Range
    {
        #region Variabili

        private int _startRow;
        private int _startColumn;
        private int _rowOffset = 0;
        private int _colOffset = 0;

        #endregion

        #region Proprietà

        public int StartRow
        {
            get { return _startRow; }
            set { _startRow = value; }
        }
        public int StartColumn
        {
            get { return _startColumn; }
            set { _startColumn = value; }
        }
        public int RowOffset
        {
            get { return _rowOffset; }
        }
        public int ColOffset
        {
            get { return _colOffset; }
        }

        #endregion

        #region Costruttori

        public Range() { }
        public Range(Range oth)
        {
            _startRow = oth.StartRow;
            _startColumn = oth.StartColumn;
            _rowOffset = oth.RowOffset;
            _colOffset = oth.ColOffset;
        }
        public Range(int row, int column)
        {
            _startRow = row;
            _startColumn = column;
        }
        public Range(int row, int column, int rowOffset)
        {
            _startRow = row;
            _startColumn = column;
            _rowOffset = rowOffset;
        }
        public Range(int row, int column, int rowOffset, int colOffset)
        {
            _startRow = row;
            _startColumn = column;
            _rowOffset = rowOffset;
            _colOffset = colOffset;
        }

        #endregion

        #region Metodi

        public void Extend(int rowOffset, int colOffset = 0) 
        {
            _rowOffset = rowOffset;
            _colOffset = colOffset;
        }
        public override string ToString()
        {
            return GetRange(_startRow, _startColumn, _rowOffset, _colOffset);
        }

        #endregion

        #region Metodi Statici

        public static string R1C1toA1(int riga, int colonna)
        {
            string output = "";
            while (colonna > 0)
            {
                int lettera = (colonna - 1) % 26;
                output = Convert.ToChar(lettera + 65) + output;
                colonna = (colonna - lettera) / 26;
            }
            output += riga;
            return output;
        }
        public static string R1C1toA1(Range cella)
        {
            return R1C1toA1(cella.StartRow, cella.StartColumn);
        }
        public static Tuple<int, int> A1toR1C1(string address)
        {
            address = address.Replace("$", "");
            string alpha = Regex.Match(address, @"\D+").Value;
            int riga = int.Parse(Regex.Match(address, @"\d+").Value);

            int colonna = 0;
            int incremento = (alpha.Length == 1 ? 1 : 26 * (alpha.Length - 1));
            for (int i = 0; i < alpha.Length; i++)
            {
                colonna += (char.ConvertToUtf32(alpha, i) - 64) * incremento;
                incremento = incremento - 26 == 0 ? 1 : incremento - 26;
            }

            return Tuple.Create<int, int>(riga, colonna);
        }
        public static string GetRange(int row, int column, int rowOffset = 0, int colOffset = 0)
        {
            if (rowOffset == 0 && colOffset == 0)
                return R1C1toA1(row, column);

            return R1C1toA1(row, column) + ":" + R1C1toA1(row + rowOffset, column + colOffset);
        }

        #endregion

    }
}
