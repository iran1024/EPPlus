using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    internal class OptimizedDependencyChain
    {
        internal List<Formula> list = new List<Formula>();
        internal Dictionary<ulong, int> index = new Dictionary<ulong, int>();
        internal List<int> CalcOrder = new List<int>();
        internal void Add(Formula f)
        {
            list.Add(f);
            //f.Index = list.Count - 1;
            //index.Add(ExcelCellBase.GetCellId(f.wsIndex, f.Row, f.Column), f.Index);
        }
    }
    internal class OptimizedDependencyChainFactory
    {
        internal static OptimizedDependencyChain Create(ExcelWorkbook wb, ExcelCalculationOption options)
        {
            var depChain = new OptimizedDependencyChain();
            foreach (var ws in wb.Worksheets)
            {
                if (ws.IsChartSheet==false)
                {
                    GetChain(depChain, wb.FormulaParser.Lexer, ws.Cells, options);
                    //GetWorksheetNames(ws, depChain, options);
                }
            }
            foreach (var name in wb.Names)
            {
                if (name.NameValue == null)
                {
                    //GetChain(depChain, wb.FormulaParser.Lexer, name, options);
                }
            }
            return depChain;
        }

        private static void GetChain(OptimizedDependencyChain depChain, ILexer lexer, ExcelRange range, ExcelCalculationOption options)
        {
            var ws = range.Worksheet;
            var fs = new CellStoreEnumerator<object>(ws._formulas, range._fromRow, range._fromCol, range._toRow, range._toCol);
            while (fs.Next())
            {
                if (fs.Value == null || fs.Value.ToString().Trim() == "") continue;
                var id = ExcelCellBase.GetCellId(ws.IndexInList, fs.Row, fs.Column);
                if (!depChain.index.ContainsKey(id))
                {
                    Formula f;
                    if (fs.Value is int ix)
                    {
                        f = ws._sharedFormulas[ix];
                    }
                    else
                    {
                        var s = fs.Value.ToString();
                        //compiler
                        if (string.IsNullOrEmpty(s)) continue;
                        f = new Formula(ws, fs.Row, fs.Column, s);
                        foreach(var address in f._addresses)
                        {

                        }
                    }
                    depChain.Add(f);
                    //FollowChain(depChain, lexer, ws.Workbook, ws, f, options);
                }
            }
        }
    }
}
