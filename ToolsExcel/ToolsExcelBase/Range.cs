using System;
using System.Collections;
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
        public Range(string range)
            : this()
        {
            Range rng = A1toRange(range);

            _startRow = rng.StartRow;
            _startColumn = rng.StartColumn;
            _rowOffset = rng.RowOffset;
            _colOffset = rng.ColOffset;
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

        public bool Contains(Range rng)
        {
            return StartRow <= rng.StartRow
                && StartColumn <= rng.StartColumn 
                && StartRow + RowOffset >= rng.StartRow + rng.RowOffset 
                && StartColumn + ColOffset >= rng.StartColumn + rng.ColOffset;
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
        public static Range A1toRange(string address)
        {
            string[] parts = address.Split(':');

            int[] rows = new int[parts.Length];
            int[] cols = new int[parts.Length];
            int j = 0;
            foreach (string part in parts)
            {
                string tmp = part.Replace("$", "");
                string alpha = Regex.Match(tmp, @"\D+").Value;
                rows[j] = int.Parse(Regex.Match(tmp, @"\d+").Value);

                cols[j] = 0;
                int incremento = (alpha.Length == 1 ? 1 : 26 * (alpha.Length - 1));
                for (int i = 0; i < alpha.Length; i++)
                {
                    cols[j] += (char.ConvertToUtf32(alpha, i) - 64) * incremento;
                    incremento = incremento - 26 == 0 ? 1 : incremento - 26;
                }
                j++;
            }

            Range rng = new Range();
            rng.StartRow = rows[0];
            rng.StartColumn = cols[0];
            
            if (rows.Length == 2)
            {
                rng.RowOffset = rows[1] - rows[0] + 1;
                rng.ColOffset = cols[1] - cols[0] + 1;
            }

            return rng;
        }
        public static string GetRange(int row, int column, int rowOffset = 1, int colOffset = 1)
        {
            if ((rowOffset == 1 && colOffset == 1))
                return R1C1toA1(row, column);

            return R1C1toA1(row, column) + ":" + R1C1toA1(row + rowOffset - 1, column + colOffset - 1);
        }

        #endregion

        #region Classi Interne

        public class RowsCollection : IEnumerable
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
            public Range this[int row1, int row2]
            {
                get
                {
                    return new Range(_r.StartRow + row1, _r.StartColumn, row2 - row1 + 1, _r.ColOffset);
                }
            }
            public int Count
            {
                get
                {
                    return _r.RowOffset;
                }
            }
            public IEnumerator GetEnumerator()
            {
                return new RowsEnum(_r);
            }
        }
        public class ColumnsCollection : IEnumerable
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
            public Range this[int col1, int col2]
            {
                get
                {
                    return new Range(_r.StartRow, _r.StartColumn + col1, _r.RowOffset, col2 - col1 + 1);
                }
            }
            public int Count
            {
                get
                {
                    return _r.ColOffset;
                }
            }
            public IEnumerator GetEnumerator()
            {
                return new ColumnsEnum(_r);
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
            public int Count
            {
                get
                {
                    return _r.ColOffset * _r.RowOffset;
                }
            }
            public IEnumerator GetEnumerator()
            {
                return new CellsEnum(_r);
            }
        }

        public class RowsEnum : IEnumerator
        {
            Range _r;
            int _position = -1;
            int _maxOffset = -1;

            public RowsEnum(Range r)
            {
                _r = r;
                _maxOffset = _r.RowOffset;
                
            }

            public object Current
            {
                get { return _r.Rows[_position]; }
            }

            public bool MoveNext()
            {
                _position++;
                return _position < _maxOffset;
            }

            public void Reset()
            {
                _position = -1;
            }
        }
        public class ColumnsEnum : IEnumerator
        {
            Range _r;
            int _position = -1;
            int _maxOffset = -1;

            public ColumnsEnum(Range r)
            {
                _r = r;
                _maxOffset = _r.ColOffset;

            }

            public object Current
            {
                get { return _r.Columns[_position]; }
            }

            public bool MoveNext()
            {
                _position++;
                return _position < _maxOffset;
            }

            public void Reset()
            {
                _position = -1;
            }
        }
        public class CellsEnum : IEnumerator
        {
            Range _r;
            int _xPosition = -1;
            int _yPosition = 0;
            int _xOffset = -1;
            int _yOffset = -1;

            public CellsEnum(Range r)
            {
                _r = r;
                _xOffset = _r.ColOffset;
                _yOffset = _r.RowOffset;

            }

            public object Current
            {
                get { return _r.Cells[_yPosition, _xPosition]; }
            }

            public bool MoveNext()
            {
                _xPosition++;
                if (_xPosition == _xOffset)
                {
                    _xPosition = 0;
                    _yPosition++;
                }
                return _yPosition < _yOffset;
            }

            public void Reset()
            {
                _xPosition = -1;
                _yPosition = 0;
            }
        }
        #endregion
    }
}
