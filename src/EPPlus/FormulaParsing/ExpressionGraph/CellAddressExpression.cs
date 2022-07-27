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
        public CellAddressExpression(Token token, ParsingContext ctx, ref FormulaAddressBase addressInfo) : base(token.Value, ctx)
        {
            if(addressInfo== null)
            {
                _addressInfo = new FormulaRangeAddress();
            }
            else
            {
                _addressInfo = new FormulaRangeAddress() { ExternalReferenceIx= addressInfo.ExternalReferenceIx, WorksheetIx=addressInfo.WorksheetIx };
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
            //Range
            _addressInfo.FromRow = _addressInfo.ToRow = row;
            _addressInfo.FromCol = _addressInfo.ToCol = col;
            _addressInfo.FixedFlag = fixedRow ? FixedFlag.FromRowFixed | FixedFlag.ToRowFixed : 0;
            _addressInfo.FixedFlag |= fixedCol ? FixedFlag.FromColFixed | FixedFlag.ToColFixed : 0;

            if ((Operator!=null && Operator.Operator == Operators.Colon) || (Prev != null && Prev.Operator.Operator == Operators.Colon))
            {

                return new AddressCompileResult(_addressInfo, DataType.ExcelCellAddress, _addressInfo);
            }
            else
            {
                // Single Cell.
                _addressInfo.WorksheetIx = _addressInfo.WorksheetIx == short.MinValue ? (short)Context.CurrentCell.WorksheetIx : _addressInfo.WorksheetIx;
                var wsIx = _addressInfo.WorksheetIx < -1 ? Context.Scopes.Current.Address.WorksheetIx : _addressInfo.WorksheetIx;
                if (wsIx < 0) return new CompileResult(eErrorType.Ref);
                var result = CompileResultFactory.Create(Context.Package.Workbook.Worksheets[wsIx].GetValueInner(row, col), 0, _addressInfo);
                if(result.IsNumeric && _negate)
                {
                    result.Negate();
                }
                return result;
            }
        }

        internal override Expression Clone()
        {
            throw new NotImplementedException();
        }
    }
}
