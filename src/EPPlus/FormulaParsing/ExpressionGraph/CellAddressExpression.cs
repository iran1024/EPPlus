using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal class CellAddressExpression : Expression
    {
        public CellAddressExpression(Token token, ParsingContext ctx) : base(token.Value, ctx)
        {
        }

        public override bool IsGroupedExpression => false;

        public bool HasCircularReference { get; internal set; }

        public override CompileResult Compile()
        {
            ExcelCellBase.GetRowColFromAddress(ExpressionString, out int row, out int col);
            var ri = Context.ExcelDataProvider.GetRange(Context.Scopes.Current.Address.Worksheet, row, col, row, col);
            return new CompileResult(ri ,DataType.ExcelCellAddress);
        }
    }
}
