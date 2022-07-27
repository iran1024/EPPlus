using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
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
            Formula f=null;
            Stack<ExpressionTree> stack=new Stack<ExpressionTree>();
            var fs = new CellStoreEnumerator<object>(ws._formulas, range._fromRow, range._fromCol, range._toRow, range._toCol);
NextFormula:
            if (fs.Next())
            {
                if (fs.Value == null || fs.Value.ToString().Trim() == "") goto NextFormula;
                var id = ExcelCellBase.GetCellId(ws.IndexInList, fs.Row, fs.Column);
                if (!depChain.index.ContainsKey(id))
                {
                    ws.Workbook.FormulaParser.ParsingContext.CurrentCell = new FormulaCellAddress(ws.IndexInList, fs.Row, fs.Column);
                    if (fs.Value is int ix)
                    {
                        f = ws._sharedFormulas[ix];
                        f.UpdateAddress(fs.Row, fs.Column);
                    }
                    else
                    {
                        var s = fs.Value.ToString();
                        //compiler
                        if (string.IsNullOrEmpty(s)) goto NextFormula;
                        f = new Formula(ws, fs.Row, fs.Column, s);
                    }
                    goto FollowFomulaChain;                    
                }
            }

        FollowFomulaChain:
            var et = f.ExpressionTree;
            if (f.AddressExpressionIndex < et.AddressExpressions.Count)
            {
                var address = et.AddressExpressions[f.AddressExpressionIndex++].Compile().Address;
                if(ws._formulas.HasValue(address.FromRow, address.FromCol, address.ToRow, address.ToCol))
                {
                    stack.Push(et);
                    if (address.FromRow==address.ToRow && address.FromCol==address.ToCol)
                    {
                        ExcelWorksheet fws;
                        if (address.WorksheetIx > 0)
                            fws = ws.Workbook.Worksheets[address.WorksheetIx];
                        else
                            fws = ws;

                        var fv = fws._formulas.GetValue(address.FromRow, address.FromCol);
                        if (fv is int ix)
                        {
                            f = fws._sharedFormulas[ix];
                            et = f.GetExpressionTree(address.FromRow, address.FromCol);
                        }
                        else
                        {
                            var s = fv.ToString();
                            //compiler
                            if (string.IsNullOrEmpty(s)) goto FollowFomulaChain;
                            f = new Formula(ws, address.FromRow, address.FromCol, s);
                        }
                        goto FollowFomulaChain;
                    }
                    else
                    {
                        f._formulaEnumerator = new CellStoreEnumerator<object>(ws._formulas, address.FromRow, address.FromCol, address.ToRow, address.ToRow);
                    }
                }
            }
            depChain.Add(f);
            if (stack.Count > 0)
            {
                et = stack.Pop();
                goto FollowFomulaChain;
            }
            goto NextFormula;
        }

        private static void FollowChain(CellStoreEnumerator<object> fs, OptimizedDependencyChain depChain, Formula f)
        {
//IterateAddresses:           
 //ProcessAddress:
            
        }
    }

}
