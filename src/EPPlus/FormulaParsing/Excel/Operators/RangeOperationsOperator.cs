using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
    internal static class RangeOperationsOperator
    {
        private static double ApplyOperator(double l, double r, Operators op)
        {
            switch(op)
            {
                case Operators.Plus:
                    return l + r;
                case Operators.Minus:
                    return l - r;
                case Operators.Multiply:
                    return l * r;
                case Operators.Divide:
                    return l / r;
                default:
                    throw new InvalidOperationException($"Operator {op} cannot be applied on ranges");
            }
        }
        public static CompileResult Apply(CompileResult left, CompileResult right, Operators op, ParsingContext context)
        {
            if(left.DataType == DataType.ExcelRange && right.DataType == DataType.ExcelRange)
            {
                var lr = left.Result as IRangeInfo;
                var rr = right.Result as IRangeInfo;

                // size of the ranges must be same...
                var sizeH = lr.Address.ToRow - lr.Address.FromRow + 1;
                if(sizeH != (rr.Address.ToRow - rr.Address.FromRow + 1))
                {
                    return new CompileResult(eErrorType.NA);
                }
                var sizeW = lr.Address.ToCol - lr.Address.FromCol + 1;
                if (sizeW != (rr.Address.ToCol - rr.Address.FromCol + 1))
                {
                    return new CompileResult(eErrorType.NA);
                }

                var scopeAdr = context.Scopes.Current.Address;
                var currentAdr = new ExcelAddress(scopeAdr.WorksheetName, scopeAdr.FromRow, scopeAdr.FromCol, scopeAdr.ToRow, scopeAdr.ToCol);
                currentAdr._toCol = currentAdr._fromCol + sizeW - 1;
                currentAdr._toRow = currentAdr._fromRow + sizeH - 1;
                var currentWs = context.Package.Workbook.Worksheets[context.Scopes.Current.Address.WorksheetName];
                var rangeAdr = new FormulaRangeAddress(context)
                {
                    WorksheetIx = (short)currentWs.PositionId,
                    FromRow = currentAdr._fromRow,
                    FromCol = currentAdr._fromCol,
                    ToRow = currentAdr._toRow,
                    ToCol = currentAdr._toCol,
                };
                var rangeDef = new RangeDefinition((short)sizeW, sizeH);
                var resultRange = new InMemoryRange(rangeAdr, rangeDef, context);
                for(var row = 0; row < sizeH; row++)
                {
                    var rowLeft = lr.Address.FromRow + row;
                    var rowRight = rr.Address.FromRow + row;
                    for(var col = 0; col < sizeW; col++)
                    {
                        var colLeft = lr.Address.FromCol + col;
                        var colRight = rr.Address.FromCol + col;

                        var leftVal = lr.GetValue(rowLeft, colLeft);
                        var rightVal = rr.GetValue(rowRight, colRight);
                        if(!ConvertUtil.IsNumeric(leftVal) || !ConvertUtil.IsNumeric(rightVal))
                        {
                            return new CompileResult(eErrorType.Value);
                        }
                        var result = ApplyOperator(ConvertUtil.GetValueDouble(leftVal), ConvertUtil.GetValueDouble(rightVal), op);
                        resultRange.SetValue(col, row, result);
                    }
                }
                return new CompileResult(resultRange, DataType.ExcelRange);
            }
            return CompileResult.Empty;
        }
    }
}
