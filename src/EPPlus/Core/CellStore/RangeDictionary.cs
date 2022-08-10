using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml;
using System;
using System.Collections.Generic;


namespace OfficeOpenXml.Core.CellStore
{
    internal class RangeDictionary
    {
        //private class RangeSorter : IComparer<FormulaRangeAddress>
        //{
        //    public int Compare(FormulaRangeAddress x, FormulaRangeAddress y)
        //    {
        //        return (x.FromCol * ExcelPackage.MaxRows + x.FromRow) - (y.ToCol * ExcelPackage.MaxRows + y.ToRow);
        //    }
        //}
        //RangeSorter _rangeSorter=new RangeSorter();
        internal Dictionary<int,List<long>> _addresses = new Dictionary<int, List<long>>();
        internal bool Merge(FormulaRangeAddress newAddress)
        {
            var isAdded = false;
            for(int c=newAddress.FromCol;c<=newAddress.ToCol;c++)
            {
                var rowSpan = ((newAddress.FromRow-1) << 20) | (newAddress.ToRow-1);
                if (!_addresses.TryGetValue(c, out List<long> rows))
                {
                    rows = new List<long>();
                    rows.Add(rowSpan);
                    _addresses.Add(c, rows);
                    continue;
                }
                var ix = rows.BinarySearch(rowSpan);
                if(ix < 0)
                {
                    long spillRange;
                    ix = ~ix;
                    if (ix > 0) ix--;
                    isAdded = VerifyAndAdd(newAddress, rowSpan, rows, ix, out spillRange);
                    if(isAdded==false && ++ix<rows.Count)
                    {
                        isAdded = VerifyAndAdd(newAddress, rowSpan, rows, ix, out spillRange);
                    }
                }
            }
            return isAdded;
        }

        private static bool VerifyAndAdd(FormulaRangeAddress newAddress, int rowSpan, List<long> rows, int ix, out long spillRange)
        {
            var fromRow = (int)(rows[ix] >> 20) + 1;
            var toRow = (int)(rows[ix] & 0xFFFFF) + 1;
            spillRange = 0;
            if (newAddress.FromRow > toRow)
            {
                if(newAddress.FromRow-1==toRow) //Next to each other: Merge
                {
                    rows[ix]= (fromRow-1 << 20) | (newAddress.ToRow-1);
                }
                else
                {
                    rows.Insert(ix + 1, rowSpan);
                }
                return true;
            }
            else if (newAddress.ToRow < fromRow)
            {
                if(newAddress.ToRow + 1 == fromRow)   //Next to each other: Merge
                {
                    rows[ix] = (newAddress.FromRow - 1 << 20) | (toRow - 1);
                }
                else
                {
                    rows.Insert(ix, rowSpan);
                }
                return true;
            }
            else
            {
                if (newAddress.FromRow >= fromRow && newAddress.ToRow <= toRow) //Within, 
                {
                    return false;
                }
                else
                {
                    if (newAddress.FromRow < fromRow && newAddress.ToRow <= toRow)
                    {
                        spillRange = ((newAddress.FromRow - 1) << 20) | (fromRow - 1);
                        rows[ix] = ((newAddress.FromRow) << 20) & (toRow - 1);
                        //Set new Span
                    }
                    if (newAddress.FromRow >= fromRow && newAddress.ToRow > toRow)
                    {
                        spillRange = (toRow << 20) & (newAddress.ToRow - 1);
                        rows[ix] = ((fromRow - 1) << 20) | (newAddress.ToRow - 1);
                        //Set new Span
                    }
                }
            }

            return false;
        }
    }
}
