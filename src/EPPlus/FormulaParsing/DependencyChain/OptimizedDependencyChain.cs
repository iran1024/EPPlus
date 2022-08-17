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
                    AddRangeToChain(depChain, wb.FormulaParser.Lexer, ws.Cells, options);
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
        internal static OptimizedDependencyChain Create(ExcelWorksheet ws, ExcelCalculationOption options)
        {
            var depChain = new OptimizedDependencyChain();

            AddRangeToChain(depChain, ws.Workbook.FormulaParser.Lexer, ws.Cells, options);

            return depChain;
        }

        private static void AddRangeToChain(OptimizedDependencyChain depChain, ILexer lexer, ExcelRange range, ExcelCalculationOption options)
        {
            var ws = range.Worksheet;
            Formula f = null;
            var fs = new CellStoreEnumerator<object>(ws._formulas, range._fromRow, range._fromCol, range._toRow, range._toCol);
            while (fs.Next())
            {

                if (fs.Value == null || fs.Value.ToString().Trim() == "") continue;
                var id = ExcelCellBase.GetCellId(ws.IndexInList, fs.Row, fs.Column);
                if (depChain.processedCells.Contains(id) == false)
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
                        if (string.IsNullOrEmpty(s)) continue;
                        f = new Formula(ws, fs.Row, fs.Column, s);
                    }
                    AddChainForFormula(depChain, lexer, f, options);
                    }
            }
        }
        private static void AddChainForFormula(OptimizedDependencyChain depChain, ILexer lexer, Formula f, ExcelCalculationOption options)
        {
            Stack<Formula> stack = new Stack<Formula>();
            var ws = f._ws;
        //            var ws = range.Worksheet;
        //            Formula f=null;
        //            ExpressionTree et=null;
        //            Stack<Formula> stack=new Stack<Formula>();            
        //            var wsFs = new CellStoreEnumerator<object>(ws._formulas, range._fromRow, range._fromCol, range._toRow, range._toCol);
        //            var fs = wsFs;
        //NextFormula:
        //            if (fs.Next())
        //            {

        //                if (fs.Value == null || fs.Value.ToString().Trim() == "") goto NextFormula;
        //                var id = ExcelCellBase.GetCellId(ws.IndexInList, fs.Row, fs.Column);
        //                if (depChain.processedCells.Contains(id)==false)
        //                {
        //                    depChain.processedCells.Add(id);
        //                    ws.Workbook.FormulaParser.ParsingContext.CurrentCell = new FormulaCellAddress(ws.IndexInList, fs.Row, fs.Column);
        //                    if (fs.Value is int ix)
        //                    {
        //                        f = ws._sharedFormulas[ix].GetFormula(fs.Row, fs.Column);
        //                    }

        //                    else
        //                    {
        //                        var s = fs.Value.ToString();
        //                        //compiler
        //                        if (string.IsNullOrEmpty(s)) goto NextFormula;
        //                        f = new Formula(ws, fs.Row, fs.Column, s);
        //                    }
        //                    goto FollowFormulaChain;                    
        //                }
        //                goto NextFormula;
        //            }
        //            else
        //            {
        //                if(fs!=wsFs)
        //                {
        //                    fs = wsFs;
        //                    f._formulaEnumerator = null;
        //                    f.AddressExpressionIndex++;
        //                    goto FollowFormulaChain;
        //                }
        //            }
        //            return;
        FollowFormulaChain:
            var et = f.ExpressionTree;
            if (f.AddressExpressionIndex < et.AddressExpressions.Count)
            {
                var address = et.AddressExpressions[f.AddressExpressionIndex++].Compile().Address;
                if (address.FromRow == address.ToRow && address.FromCol == address.ToCol)
                {
                    if (GetProcessedAddress(depChain, (int)address.WorksheetIx, address.FromRow, address.FromCol) && 
                        ws._formulas.Exists(address.FromRow, address.FromCol))
                    {
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
                    f._formulaEnumerator = new CellStoreEnumerator<object>(ws._formulas, address.FromRow, address.FromCol, address.ToRow, address.ToCol);
                    goto NextFormula;
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
                if (f._formulaEnumerator == null)
                {
                    goto FollowFormulaChain;
                }
                else
                {
                    goto NextFormula;
                }
            }
            return;
        NextFormula:
            var fs = f._formulaEnumerator;
            if (f._formulaEnumerator.Next())
            {
                if (fs.Value == null || fs.Value.ToString().Trim() == "") goto NextFormula;
                var id = ExcelCellBase.GetCellId(ws.IndexInList, fs.Row, fs.Column);
                if (depChain.processedCells.Contains(id) == false)
                {
                    depChain.processedCells.Add(id);
                    ws.Workbook.FormulaParser.ParsingContext.CurrentCell = new FormulaCellAddress(ws.IndexInList, fs.Row, fs.Column);
                    stack.Push(f);
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
            f._formulaEnumerator = null;
            goto FollowFormulaChain;
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
