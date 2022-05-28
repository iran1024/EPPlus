using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    internal struct RangeDefinition
    {
        public RangeDefinition(short nCols, int nRows)
        {
            NumberOfCols = nCols;
            NumberOfRows = nRows;
        }

        public RangeDefinition(short fromCol, int fromRow, short toCol, int toRow)
        {
            NumberOfCols = (short)(toCol - fromCol);
            NumberOfRows = toRow - fromRow;
        }

        public short NumberOfCols { get; private set; }

        public int NumberOfRows { get; private set; }
    }
}
