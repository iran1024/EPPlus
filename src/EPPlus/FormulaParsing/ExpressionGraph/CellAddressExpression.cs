using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Diagnostics;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    [DebuggerDisplay("CellAddressExpression: {ExpressionString}")]
    internal class CellAddressExpression : Expression
    {
        FormulaRangeAddress _addressInfo;
        public CellAddressExpression(Token token, ParsingContext ctx, FormulaRangeAddress addressInfo) : base(token.Value, ctx)
        {
            _addressInfo = addressInfo;
        }
        public override bool IsGroupedExpression => false;

        public bool HasCircularReference { get; internal set; }

        public override CompileResult Compile()
        {
            ExcelCellBase.GetRowColFromAddress(ExpressionString, out int row, out int col);
            _addressInfo.FromRow = row;
            _addressInfo.FromCol = col;
            _addressInfo.ToRow = row;
            _addressInfo.ToCol = col;

            var ri = Context.ExcelDataProvider.GetRange(_addressInfo);
            return new CompileResult(ri ,DataType.ExcelCellAddress);
        }
    }
}
