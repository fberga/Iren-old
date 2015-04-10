using System;
using System.Text.RegularExpressions;

namespace Iren.ToolsExcel.Base
{
    public class Range
    {
        #region Variabili

        private int _startRow;
        private int _startColumn;
        private int _rowOffset = 1;
        private int _colOffset = 1;

        private RowsCollection _rows;
        private ColumnsCollection _cols;
        private CellsCollection _cells;

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
            set { _rowOffset = value < 1 ? 1 : value; }
        }
        public int ColOffset
        {
            get { return _colOffset; }
            set { _colOffset = value < 1 ? 1 : value; }
        }

        public RowsCollection Rows
        {
            get { return _rows; }
        }
        public ColumnsCollection Columns
        {
            get { return _cols; }
        }
        public CellsCollection Cells
        {
            get { return _cells; }
        }

        #endregion

        #region Costruttori

        public Range() 
        {
            _rows = new RowsCollection(this);
            _cols = new ColumnsCollection(this);
            _cells = new CellsCollection(this);
        }
        public Range(Range oth) 
            : this()
        {
            _startRow = oth.StartRow;
            _startColumn = oth.StartColumn;
            _rowOffset = oth.RowOffset;
            _colOffset = oth.ColOffset;
        }
        public Range(int row, int column)
            : this()
        {
            _startRow = row;
            _startColumn = column;
        }
        public Range(int row, int column, int rowOffset)
            : this()
        {
            _startRow = row;
            _startColumn = column;
            _rowOffset = rowOffset;
        }
        public Range(int row, int column, int rowOffset, int colOffset)
            : this()
        {
            _startRow = row;
            _startColumn = column;
            _rowOffset = rowOffset;
            _colOffset = colOffset;
        }

        #endregion

        #region Metodi

        public Range Extend(int rowOffset = 1, int colOffset = 1) 
        {
            RowOffset = rowOffset;
            ColOffset = colOffset;
            
            return this;
        }
        public Range ExtendOf(int rowOffset = 0, int colOffset = 0)
        {
            _rowOffset += rowOffset;
            _colOffset += colOffset;

            return this;
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
        public static string GetRange(int row, int column, int rowOffset = 1, int colOffset = 1)
        {
            if ((rowOffset == 1 && colOffset == 1))
                return R1C1toA1(row, column);

            return R1C1toA1(row, column) + ":" + R1C1toA1(row + rowOffset - 1, column + colOffset - 1);
        }

        #endregion

        #region Classi Interne

        public class RowsCollection
        {
            private Range _r;

            internal RowsCollection(Range r)
            {
                _r = r;
            }

            public Range this[int row]
            {
                get
                {
                    return new Range(_r.StartRow + row, _r.StartColumn, 1, _r.ColOffset);
                }
            }
            public Range this[int row1, int offset]
            {
                get
                {
                    return new Range(_r.StartRow + row1, _r.StartColumn, offset - row1, _r.ColOffset);
                }
            }
            public int Count
            {
                get
                {
                    return _r.RowOffset;
                }
            }
        }
        public class ColumnsCollection
        {
            private Range _r;

            internal ColumnsCollection(Range r)
            {
                _r = r;
            }

            public Range this[int column]
            {
                get
                {
                    return new Range(_r.StartRow, _r.StartColumn + column, _r.RowOffset, 1);
                }
            }
            public Range this[int col1, int offset]
            {
                get
                {
                    return new Range(_r.StartRow, _r.StartColumn + col1, _r.RowOffset, offset - col1);
                }
            }
            public int Count
            {
                get
                {
                    return _r.ColOffset;
                }
            }
        }
        public class CellsCollection
        {
            private Range _r;

            internal CellsCollection(Range r)
            {
                _r = r;
            }

            public Range this[int row, int column]
            {
                get
                {
                    return new Range(_r.StartRow + row, _r.StartColumn + column);
                }
            }
        }

        #endregion

    }
}
