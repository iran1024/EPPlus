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
        internal List<Formula> formulas = new List<Formula>();
        internal Dictionary<int, RangeDictionary> accessedRanges = new Dictionary<int, RangeDictionary>();
        internal HashSet<ulong> processedCells = new HashSet<ulong>();
        internal void Add(Formula f)
        {
            formulas.Add(f);
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
            Stack<Formula> stack=new Stack<Formula>();            
            var fs = new CellStoreEnumerator<object>(ws._formulas, range._fromRow, range._fromCol, range._toRow, range._toCol);
NextFormula:
            if (fs.Next())
            {
                if (fs.Value == null || fs.Value.ToString().Trim() == "") goto NextFormula;
                var id = ExcelCellBase.GetCellId(ws.IndexInList, fs.Row, fs.Column);
                if (depChain.processedCells.Contains(id)==false)
                {
                    depChain.processedCells.Add(id);
                    ws.Workbook.FormulaParser.ParsingContext.CurrentCell = new FormulaCellAddress(ws.IndexInList, fs.Row, fs.Column);
                    if (fs.Value is int ix)
                    {
                        f = ws._sharedFormulas[ix].GetFormula(fs.Row, fs.Column);
                    }

                    else
                    {
                        var s = fs.Value.ToString();
                        //compiler
                        if (string.IsNullOrEmpty(s)) goto NextFormula;
                        f = new Formula(ws, fs.Row, fs.Column, s);
                    }
                    goto FollowFormulaChain;                    
                }
                goto NextFormula;
            }
            return;
FollowFormulaChain:
            var et = f.ExpressionTree;
            if (f.AddressExpressionIndex < et.AddressExpressions.Count)
            {
                var address = et.AddressExpressions[f.AddressExpressionIndex++].Compile().Address;
                if (address.FromRow == address.ToRow && address.FromCol == address.ToCol && GetProcessedAddress(depChain, (int)address.WorksheetIx, address.FromRow, address.FromCol))
                {
                    if (ws._formulas.Exists(address.FromRow, address.FromCol))
                    {
                        //var key = ExcelCellBase.GetCellId(address.WorksheetIx, address.FromRow, address.FromCol);
                        //if (depChain.index.ContainsKey(key) == false)
                        //{
                        stack.Push(f);
                        ExcelWorksheet fws;
                        if (address.WorksheetIx > 0)
                            fws = ws.Workbook.Worksheets[address.WorksheetIx];
                        else
                            fws = ws;

                        var fv = fws._formulas.GetValue(address.FromRow, address.FromCol);
                        if (fv is int ix)
                        {
                            f = fws._sharedFormulas[ix].GetFormula(address.FromRow, address.FromCol);
                        }
                        else
                        {
                            var s = fv.ToString();
                            //compiler
                            if (string.IsNullOrEmpty(s)) goto FollowFormulaChain;
                            f = new Formula(ws, address.FromRow, address.FromCol, s);
                        }
                        depChain.processedCells.Add(ExcelCellBase.GetCellId(f._ws.IndexInList, f.StartRow, f.StartCol));
                        goto FollowFormulaChain;
                    }
                }
                else if (GetProcessedAddress(depChain, ref address))
                {
                    f._formulaEnumerator = new CellStoreEnumerator<object>(ws._formulas, address.FromRow, address.FromCol, address.ToRow, address.ToRow);
                }
                if (f.AddressExpressionIndex < et.AddressExpressions.Count)
                {
                    f.AddressExpressionIndex++;
                    goto FollowFormulaChain;
                }
            }
            depChain.Add(f);
            if (stack.Count > 0)
            {
                f = stack.Pop();
                f.AddressExpressionIndex++;
                goto FollowFormulaChain;
            }
            goto NextFormula;
        }

        private static void FollowChain(CellStoreEnumerator<object> fs, OptimizedDependencyChain depChain, Formula f)
        {
//IterateAddresses:           
 //ProcessAddress:
            
        }
        private static bool GetProcessedAddress(OptimizedDependencyChain depChain, ref FormulaRangeAddress address)
        {
            if (depChain.accessedRanges.TryGetValue(address.WorksheetIx, out RangeDictionary wsRd) == false)
            {
                wsRd = new RangeDictionary();
                depChain.accessedRanges.Add(address.WorksheetIx, wsRd);
            }
            return wsRd.Merge(ref address);
        }
        private static bool GetProcessedAddress(OptimizedDependencyChain depChain, int wsIndex, int row, int col)
        {
            if (depChain.accessedRanges.TryGetValue(wsIndex, out RangeDictionary wsRd) == false)
            {
                wsRd = new RangeDictionary();
                depChain.accessedRanges.Add(wsIndex, wsRd);
            }
            return wsRd.Merge(row, col);
        }
    }
}
