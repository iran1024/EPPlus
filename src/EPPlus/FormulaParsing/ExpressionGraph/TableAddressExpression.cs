using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Diagnostics;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    [DebuggerDisplay("TableAddressExpression: {_addressInfo}")]
    internal class TableAddressExpression : Expression
    {
        readonly FormulaRangeAddress _addressInfo;
        public TableAddressExpression(ParsingContext ctx, FormulaRangeAddress addressInfo) : base(null, ctx)
        {
            _addressInfo = addressInfo;
        }
        public override bool IsGroupedExpression => false;
        public bool HasCircularReference { get; internal set; }

        internal override ExpressionType ExpressionType => ExpressionType.TableAddress;

        public override CompileResult Compile()
        {
            var ri = Context.ExcelDataProvider.GetRange(_addressInfo);
            return new CompileResult(ri, DataType.ExcelRange);
        }

        internal override Expression Clone()
        {
            return CloneMe(new TableAddressExpression(Context, _addressInfo));
        }
    }
}
