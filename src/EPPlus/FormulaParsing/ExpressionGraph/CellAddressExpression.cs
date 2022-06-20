using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Diagnostics;
using Operators = OfficeOpenXml.FormulaParsing.Excel.Operators.Operators;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    [DebuggerDisplay("CellAddressExpression: {ExpressionString}")]
    internal class CellAddressExpression : Expression
    {
        FormulaCellAddress _addressInfo;
        bool _negate;
        public CellAddressExpression(Token token, ParsingContext ctx, ref FormulaAddressBase addressInfo) : base(token.Value, ctx)
        {
            if(addressInfo== null)
            {
                _addressInfo = new FormulaCellAddress();
            }
            else
            {
                _addressInfo = new FormulaCellAddress(addressInfo);
            }
            addressInfo = _addressInfo;
            _negate = token.IsNegated;
        }
        public override bool IsGroupedExpression => false;

        public bool HasCircularReference { get; internal set; }

        internal override ExpressionType ExpressionType => ExpressionType.CellAddress;

        public override CompileResult Compile()
        {
            ExcelCellBase.GetRowColFromAddress(ExpressionString, out int row, out int col, out bool fixedRow, out bool fixedCol);
            if((Operator!=null && Operator.Operator == Operators.Colon) || (Prev != null && Prev.Operator.Operator == Operators.Colon))
            {
                //Range
                _addressInfo.WorksheetIx = _addressInfo.WorksheetIx == short.MinValue ? Context.Scopes.Current.Address.WorksheetIx : _addressInfo.WorksheetIx;
                _addressInfo.Row = row;
                _addressInfo.Col = col;
                _addressInfo.FixedRow = fixedRow;
                _addressInfo.FixedCol |= fixedCol;

                return new CompileResult(_addressInfo, DataType.ExcelCellAddress);
            }
            else
            {
                // Single Cell.
                var wsIx = _addressInfo.WorksheetIx < -1 ? Context.Scopes.Current.Address.WorksheetIx : _addressInfo.WorksheetIx;
                if (wsIx < 0) return new CompileResult(eErrorType.Ref);
                var result = CompileResultFactory.Create(Context.Package.Workbook.Worksheets[wsIx].GetValueInner(row, col));
                if(result.IsNumeric && _negate)
                {
                    result.Negate();
                }
                return result;
            }
        }
    }
}
