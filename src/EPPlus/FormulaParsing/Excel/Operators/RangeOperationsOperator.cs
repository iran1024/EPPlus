using OfficeOpenXml.FormulaParsing.ExpressionGraph;
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
                var sizeH = lr.Address._toRow - lr.Address._fromCol + 1;
                if(sizeH != (rr.Address._toRow - rr.Address._fromRow + 1))
                {
                    return new CompileResult(eErrorType.NA);
                }
                var sizeW = lr.Address._toCol - lr.Address._fromCol + 1;
                if (sizeW != (rr.Address._toCol - rr.Address._fromCol + 1))
                {
                    return new CompileResult(eErrorType.NA);
                }

                var currentAdr = new ExcelAddress(context.Scopes.Current.Address.Worksheet, context.Scopes.Current.Address.Address);
                currentAdr._toCol = currentAdr._fromCol + sizeW - 1;
                currentAdr._toRow = currentAdr._fromRow + sizeH - 1;
                var currentWs = context.Package.Workbook.Worksheets[context.Scopes.Current.Address.Worksheet];
                var resultRange = new InMemoryRange(currentWs.PositionId, currentAdr._fromRow, currentAdr._fromCol, currentAdr._toRow, currentAdr._toCol, context);
                for(var row = 0; row < sizeH; row++)
                {
                    var rowLeft = lr.Address._fromRow + row;
                    var rowRight = rr.Address._fromRow + row;
                    for(var col = 0; col < sizeW; col++)
                    {
                        var colLeft = lr.Address._fromCol + col;
                        var colRight = rr.Address._fromCol + col;

                        var leftVal = lr.GetValue(rowLeft, colLeft);
                        var rightVal = rr.GetValue(rowRight, colRight);
                        if(!ConvertUtil.IsNumeric(leftVal) || !ConvertUtil.IsNumeric(rightVal))
                        {
                            return new CompileResult(eErrorType.Value);
                        }
                        var result = ApplyOperator(ConvertUtil.GetValueDouble(leftVal), ConvertUtil.GetValueDouble(rightVal), op);
                        resultRange.SetValue(row, col, result);
                    }
                }
                return new CompileResult(resultRange, DataType.ExcelRange);
            }
            return CompileResult.Empty;
        }
    }
}
