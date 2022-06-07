using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Diagnostics;
using Operators = OfficeOpenXml.FormulaParsing.Excel.Operators.Operators;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    [DebuggerDisplay("CellAddressExpression: {ExpressionString}")]
    internal class CellAddressExpression : Expression
    {
        FormulaRangeAddress _addressInfo;
        bool _negate;
        public CellAddressExpression(Token token, ParsingContext ctx, FormulaRangeAddress addressInfo) : base(token.Value, ctx)
        {
            if(addressInfo== null)
            {
                _addressInfo = new FormulaRangeAddress(ctx);
            }
            else
            {
                _addressInfo = addressInfo;
            }
            _negate = token.IsNegated;
        }
        public override bool IsGroupedExpression => false;

        public bool HasCircularReference { get; internal set; }

        public override CompileResult Compile()
        {
            ExcelCellBase.GetRowColFromAddress(ExpressionString, out int row, out int col, out bool fixedRow, out bool fixedCol);
            //if((Operator!=null && Operator.Operator == Operators.Colon) || (Prev != null && Prev.Operator.Operator == Operators.Colon) || )
            //{
                // Range
                _addressInfo.FromRow = row;
                _addressInfo.FromCol = col;
                _addressInfo.ToRow = row;
                _addressInfo.ToCol = col;
                _addressInfo.FixedFlag = fixedRow ? FixedFlag.FromRowFixed | FixedFlag.ToRowFixed : 0;
                _addressInfo.FixedFlag |= fixedCol ? FixedFlag.FromColFixed | FixedFlag.ToColFixed : 0;
                var ri = Context.ExcelDataProvider.GetRange(_addressInfo);

                return new CompileResult(ri, DataType.ExcelCellAddress);
            //}
            //else
            //{
            //    // Single Cell.
            //    var wsIx = _addressInfo.WorksheetIx < -1 ? Context.Scopes.Current.Address.WorksheetIx : _addressInfo.WorksheetIx;
            //    if (wsIx < 0) return new CompileResult(eErrorType.Ref);
            //    var result = CompileResultFactory.Create(Context.Package.Workbook.Worksheets[wsIx].GetValueInner(row, col));
            //    if(result.IsNumeric && _negate)
            //    {
            //        result.Negate();
            //    }
            //    return result;
            //}
        }
    }
}
