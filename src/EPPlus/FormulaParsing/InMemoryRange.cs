using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    internal class InMemoryRange : IRangeInfo
    {
        public InMemoryRange(int worksheetIx, int fromRow, int fromCol, int toRow, int toCol, ParsingContext ctx)
        {
            _ws = ctx.Package.Workbook.Worksheets[ctx.Scopes.Current.Address.Worksheet];
            _address = new ExcelAddressBase(_ws.Name, fromRow, fromCol, toRow, toCol);
            _rangeNew = new FormulaRangeAddress(ctx);
            _rangeNew.FromRow = fromRow;
            _rangeNew.ToRow = toRow;
            _rangeNew.FromCol = fromCol;
            _rangeNew.ToCol = toCol;
            _nRows = toRow - fromRow + 1;
            _nCols = toCol - fromCol + 1;
            _cells = new ICellInfo[_nRows, _nCols];
        }

        private readonly ExcelAddressBase _address;
        private readonly FormulaRangeAddress _rangeNew;
        private readonly ExcelWorksheet _ws;
        private readonly ICellInfo[,] _cells;
        private int _colIx = 0;
        private int _rowIndex = -1;
        private readonly int _nRows, _nCols;

        public void SetValue(int row, int col, object val)
        {
            var c = new InMemoryCellInfo(val);
            _cells[row, col] = c;
        }

        public bool IsEmpty => _cells.Length == 0;

        public bool IsMulti => _cells.Length > 0 && _cells.GetUpperBound(1) > 1;

        public bool IsInMemoryRange => true;

        public ExcelAddressBase Address => _address;

        public FormulaRangeAddress RangeNew => _rangeNew;

        public ExcelWorksheet Worksheet => _ws;

        public ICellInfo Current => _cells[_rowIndex, _colIx];

        object IEnumerator.Current => _cells[_rowIndex, _colIx];

        public void Dispose()
        {
            
        }

        public IEnumerator<ICellInfo> GetEnumerator()
        {
            return this;
        }

        public int GetNCells()
        {
            return _cells.Length * _cells.GetUpperBound(1);
        }

        public object GetOffset(int rowOffset, int colOffset)
        {
            var c = _cells[rowOffset, colOffset];
            if(c == null)
            {
                return null;
            }
            return c.Value;
        }

        public object GetValue(int row, int col)
        {
            var c = _cells[row, col];
            if (c == null) return null;
            return c.Value;
        }

        public bool MoveNext()
        {
            if(_colIx < _nCols - 1)
            {
                _colIx++;
                return true;
            }
            _colIx = 0;
            _rowIndex++;
            if (_rowIndex >= _nRows) return false;
            return true;
        }

        public void Reset()
        {
            _colIx = 0;
            _rowIndex = 0;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this;
        }
    }
}
